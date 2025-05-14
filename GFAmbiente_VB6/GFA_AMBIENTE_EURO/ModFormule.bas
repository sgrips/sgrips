Attribute VB_Name = "ModFormule"
Option Explicit

Public Enum TpoCoeff
    CoeffKA
    CoeffKB
    CoeffKC
    CoeffKD
    CoeffUA507
    CoeffUD507
End Enum

Public Enum TpoBonus
    BonusUtenzaAttivita
    BonusUtenzaDomestica
End Enum

Public Enum TpoContenitoriSvuotaInDotazione
    ContenitoriSingoliSvuotamentiUD
    ContenitoriSingoliSvuotamentiUA
    ContenitoriSvuotamentiUD
    ContenitoriSvuotamentiUA
    ContenitoriInDotazioneUA
    ContenitoriInDotazioneUD
End Enum



Public Enum tpoTariffa
    TariffaNonImpostata = -1
    TariffaNormalizzata = 0
    Tariffa507 = 1
    TariffaPortaAPorta = 3
    TariffaTransitorio = 2
    TariffaCostiIndivisibili = 4
End Enum


Function InEuro(X As Double) As Double
    'Converte da Lire a €
    If X = 0 Then Exit Function
    InEuro = X / CDbl(GetParamValue("1900", vbNullString, PARAM_EURO))
End Function

Function InLire(X As Double) As Double
    'Converte da € a Lire
    If X = 0 Then Exit Function
    InLire = X * (CDbl(GetParamValue("1900", vbNullString, PARAM_EURO)))
End Function


Function CalcoloET1Transitorio(ET0 As Double, IP As Double, X As Double, Y As Double, CCON As Double, K As Double) As Double
    'Calcolo del Totale entrate tariffare attese
    'IP , X , Y , CCON , K sono espressi come persentuali
    CalcoloET1Transitorio = ET0 * (1 + IP / 100 - X / 100 - Y / 100 - CCON / 100 + K / 100)
End Function



Function CalcoloET1Integrale(ByVal IP As Double, ByVal RP As Double, ByVal CG As Double, ByVal CC As Double, _
        ByVal CK As Double, ByVal CT As Double, ETF As Double, ETV As Double) As Double
    
    Dim Alfa As Double
    'Calcolo del Totale entrate tariffare attese
    'IP , RP sono espressi come persentuali
    'Input/Output
    '   ETF = Parte Fissa
    '   ETV = Parte Variabile
    'OutPut
    '   ET  = Totale Entrate Tariffarie di riferimento
    '   ETF = Parte Fissa Rivisitata
    '   ETV = Parte Variabile Rivisitata
    CalcoloET1Integrale = (CG + CC) * (1 + IP / 100 - RP / 100) + CK
    
    If CT = 0 Then Exit Function
    
    Alfa = CT / CalcoloET1Integrale
    ETF = ETF / Alfa
    ETV = ETV / Alfa
    
End Function

Function DatiCostiComune(ByVal CodComune As String, ByVal AnnoRiferimento As String) As Recordset
    'Reperisco i Costi del comune
    Set DatiCostiComune = gDBC.Execute("Select CostoComunale.*," _
    & "CSL + CARC + CGG + CCD + AC + CK as ParteFissa," _
    & "CRT + CTS + CRD + CTR - CON - Y as ParteVariabile " _
    & "from CostoComunale  where AnnoRiferimento='" & AnnoRiferimento & "' and CodComune='" & CodComune & "'")
    
End Function
    
Function DatiTecniciComune(ByVal CodComune As String, ByVal AnnoRiferimento As String) As Recordset
    Dim rs As Recordset
        
    'Dati Tecnici del comune
    Set rs = DatiCostiComune(CodComune, AnnoRiferimento)
    If Not rs.EOF Then
        Set DatiTecniciComune = gDBC.Execute("Select DatiTecniciComune.*, NRuoliSecondaCasa +NRuoliUtenzaDomestica as NruoliUtenzaD ," _
        & "QRifiutiTotali - QRifiutiCategorieAttivita as  QRifiutiUtenzaDomestica," _
        & "QRifiutiTotali - QRifiutiRaccoltaDifferenziata as QRifiutiSmaltiti," _
        & "NRuoliUtenzaDomestica +NRuoliSecondaCasa + NRuoliUtenzaAttivita as NRuoliTot, " _
        & "TotSuperficieUtenzaDomestica + TotSuperficieUtenzaAttivita as TotSuperficieIscrRuolo," _
        & Replace(rs("Tot_Servizio"), ",", ".") & "/QRifiutiTotali as CTotKgRifiutiGestiti," _
        & Replace(rs("CRD"), ",", ".") & "/ QRifiutiRaccoltaDifferenziata  as CKgRifiutiRaccDiff," _
        & Replace(rs("ParteFissa"), ",", ".") & "/QRifiutiTotali as CTotKgRifiutiQF," _
        & Replace(rs("ParteVariabile"), ",", ".") & "/QRifiutiTotali as CTotKgRifiutiQV," _
        & 0 & "/1 as CTotFissiUD " _
        & "from DatiTecniciComune  where AnnoRiferimento='" & AnnoRiferimento & "' and CodComune='" & CodComune & "'")
    Else
        Set DatiTecniciComune = rs
    End If
    
End Function

Function DatiStatisticiComune(ByVal CodComune As String, ByVal AnnoRiferimento As String) As Recordset

        
    'Dati Tecnici del comune
    Set DatiStatisticiComune = gDBC.Execute("Select * from Statistiche ST inner join StatisticheComunali  STC on  ST.CodStatistica=STC.CodStatistica Where STC.CodComune is null or (STC.CodComune='" & CodComune & "' and STC.AnnoRiferimento='" & AnnoRiferimento & "') order by ST.CodStatistica")
    
       
End Function

Function SelezionaTotaliSvuotamentiPerContenitore(ByVal CodComune As String, ByVal AnnoRiferimento As String) As Recordset
    Dim qry As Command

    Set qry = gADOX.Procedures("SelSvuotamentiPerContenitore").Command
        
    qry.Parameters.Refresh
    qry.Parameters("pCodComune") = CodComune
    qry.Parameters("pAnnoRiferimento") = AnnoRiferimento

    Set SelezionaTotaliSvuotamentiPerContenitore = qry.Execute
    
End Function

Function DatiFinaliTariffa(ByVal CodComune As String, ByVal AnnoRiferimento As String, TipoTariffa As tpoTariffa) As Recordset
    Dim rs As Recordset
    Dim qry As Command
    
    Select Case TipoTariffa
        Case TariffaNormalizzata, TariffaCostiIndivisibili
            Set rs = gDBC.Execute("Select * from DatiFinaliCalcoloTariffa " & _
                 "where AnnoRiferimento='" & AnnoRiferimento & "' and CodComune='" & CodComune & "'")
            Set DatiFinaliTariffa = rs
        Case Tariffa507
            Set rs = gDBC.Execute("Select * from DatiFinaliCalcoloTarsu507 " & _
                 "where AnnoRiferimento='" & AnnoRiferimento & "' and CodComune='" & CodComune & "'")
            Set DatiFinaliTariffa = rs
    End Select

End Function



Function DatiAnagraficiComunali(ByVal mCodComune As String, ByVal mAnno As String, Optional ParAbitanti As String) As Recordset
    Dim rs As Recordset
    
    'reperisco i dati del comune
    Set rs = gDBC.Execute("SELECT Comuni.Comune,  Provincie.DESCRIZION, Regioni.REGIONE, ParametriSistema.Parametro as AreaGeografica, ParametriSistema.Descrizione ,AnagraficaComunale.* " & _
                                "FROM AnagraficaComunale INNER JOIN ((Comuni INNER JOIN (Provincie INNER JOIN Regioni ON Provincie.CodRegione = Regioni.CodRegione) ON Comuni.Provincia = Provincie.PROVINCIA) " & _
                                "INNER JOIN ParametriSistema ON Regioni.CodArea = ParametriSistema.Parametro) ON AnagraficaComunale.CodComune = Comuni.CODISTAT WHERE Comuni.CodIstat='" & mCodComune & "' and AnagraficaComunale.AnnoRiferimento='" & mAnno & "'")
        
            
    'Area di appartenenza del comune , NumeroAbitanti
    Set DatiAnagraficiComunali = rs
    
    If rs.EOF Then Exit Function
    
    If rs("NumAbitantiResidenti") <= CDbl(GetParamValue(mAnno, mCodComune, "DCAMIN")) Then
        ParAbitanti = "DCAMIN"
    Else
        ParAbitanti = "DCAMAG"
    End If
End Function



Function QuantitaRifiutiDaUtenzeAttivita(ByVal mCodComune As String, ByVal mAnnoRiferimento As String) As Double
    Dim rs As Recordset
    Dim rsAnaComunale As Recordset
    Dim qry As ADODB.Command
    Dim ParAbitanti As String
    Dim QRifCA As Double
    
    QuantitaRifiutiDaUtenzeAttivita = 0
        
    Set rsAnaComunale = DatiAnagraficiComunali(mCodComune, mAnnoRiferimento, ParAbitanti)
    
    If Not rsAnaComunale.EOF Then
    
        'Seleziono i KD
        Set rs = SelezionaCoeffK(CoeffKD, mCodComune, mAnnoRiferimento, rsAnaComunale("AreaGeografica"), ParAbitanti)
        
        'Calcolo quantità rifiuti
        QRifCA = 0
        Do While Not rs.EOF
            QRifCA = QRifCA + (rs("Valore") * rs("Superficie"))
            rs.MoveNext
        Loop
        
    End If
    
    QuantitaRifiutiDaUtenzeAttivita = QRifCA
    
    'Update Quantita Rifiuti Da UA su Tabella DatiTecnici
    gDBC.Execute "Update DatiTecniciComune set QRifiutiCategorieAttivita='" & QRifCA & "'  Where AnnoRiferimento='" & mAnnoRiferimento & "' and CodComune='" & mCodComune & "'"
End Function


Function NumeroUtenzeDomesticheCorreteKB(ByVal mCodComune As String, ByVal mAnnoRiferimento As String) As Double
    Dim rs As Recordset
    Dim rsDatiAnagraficiComune As Recordset
    Dim qry As Command
    Dim ParAbitanti As String
    Dim Atot As Double
    Dim NComFamiglia As Double
    Dim NComFamigliaR As Double
    Dim RsDAtiTecComune As Recordset
    
    NumeroUtenzeDomesticheCorreteKB = 0
        
    Set rsDatiAnagraficiComune = DatiAnagraficiComunali(mCodComune, mAnnoRiferimento, ParAbitanti)
    Set RsDAtiTecComune = DatiTecniciComune(mCodComune, mAnnoRiferimento)
    
    If Not rsDatiAnagraficiComune.EOF Then
    
        'KB
        Set rs = SelezionaCoeffK(CoeffKB, mCodComune, mAnnoRiferimento, 0, 0)
        
        
        'Calcolo
        Atot = 0
        Do While Not rs.EOF
        
            'Numero Componenti
            NComFamiglia = 0
            Select Case rs("IDNumeroComponenti")
                Case 1
                    NComFamiglia = rsDatiAnagraficiComune("FamigliaCon" & CStr(rs("IDNumeroComponenti")) & "Componente")
                Case Is >= 2
                    If rs("IDNumeroComponenti") = 6 Then
                        NComFamiglia = rsDatiAnagraficiComune("FamigliaConPiuDi5Componenti")
                    Else
                        NComFamiglia = rsDatiAnagraficiComune("FamigliaCon" & CStr(rs("IDNumeroComponenti")) & "Componenti")
                    End If
            End Select
            'Aggiundo i numero ruoli seconda casa se vi sono
            If RsDAtiTecComune("Id_NumCompSecondaCasa") = rs("IDNumeroComponenti") Then
                NComFamiglia = NComFamiglia + RsDAtiTecComune("NRuoliSecondaCasa")
            End If
            
            Atot = Atot + (rs("Valore") * (NComFamiglia))
            rs.MoveNext
        Loop
        
    End If
    
    NumeroUtenzeDomesticheCorreteKB = Atot
    
End Function

Function SuperficeCorrettaUD(ByVal mCodComune As String, ByVal mAnnoRiferimento As String, SupMedieCalc() As Double, SupCorrettaSenzaRiduzione() As Double, rsRIDUD As ADODB.Recordset, ConRiduzione As Boolean) As Boolean
    Dim rs As Recordset
    Dim rsDatiAnagraficiComune As Recordset
    Dim qry As Command
    Dim ParAbitanti As String
    Dim TotSupMedieCalc As Double
    Dim NComFamiglia As Long
    Dim RsDAtiTecComune As Recordset
    Dim RsDAtiUD As Recordset
    Dim SuperficieMedia As Double
    Dim i As Integer
    Dim Alfa As Double
    Dim NComFamiglie(6) As Double
    Dim TotSuperficieInRiduzioniDAUD_PF As Double
    
    
    SuperficeCorrettaUD = False
    'Se la superficie totale relativa al n. componenti viene fornita dal comune non occorre calcolare quella
    'sperimentale
    Set RsDAtiUD = gDBC.Execute("select * from UtenzeDomesticheComunali where CodComune='" & mCodComune & "' and AnnoRiferimento='" & mAnnoRiferimento & "' order by IdNumeriComponenti")
    Set rsDatiAnagraficiComune = DatiAnagraficiComunali(mCodComune, mAnnoRiferimento, ParAbitanti)
    
    
    If IsNull(RsDAtiUD("Superficie")) Or RsDAtiUD("Superficie") = 0 Then
        Set RsDAtiTecComune = DatiTecniciComune(mCodComune, mAnnoRiferimento)
        
        If Not rsDatiAnagraficiComune.EOF And Not RsDAtiTecComune.EOF Then
        
            SuperficieMedia = RsDAtiTecComune("TotSuperficieUtenzaDomestica") / RsDAtiTecComune("NruoliUtenzaD")
        
            'KA
            Set rs = SelezionaCoeffK(CoeffKA, mCodComune, mAnnoRiferimento, rsDatiAnagraficiComune("AreaGeografica"), ParAbitanti)
            
            If rs.EOF Then Exit Function
            
            'Calcolo
            Do While Not rs.EOF
            
                i = -1
                On Error Resume Next
                i = UBound(SupMedieCalc)
                On Error GoTo 0
                ReDim Preserve SupMedieCalc(i + 1)
                
                'Numero Componenti
                Select Case rs("IDNumeroComponenti")
                    Case 1
                        NComFamiglia = rsDatiAnagraficiComune("FamigliaCon" & CStr(rs("IDNumeroComponenti")) & "Componente")
                         
                    Case Is >= 2
                        If rs("IDNumeroComponenti") = 6 Then
                            NComFamiglia = rsDatiAnagraficiComune("FamigliaConPiuDi5Componenti")
                        Else
                            NComFamiglia = rsDatiAnagraficiComune("FamigliaCon" & CStr(rs("IDNumeroComponenti")) & "Componenti")
                        End If
                End Select
    
                'Aggiundo i numero ruoli seconda casa se vi sono
                If RsDAtiTecComune("Id_NumCompSecondaCasa") = rs("IDNumeroComponenti") Then
                    NComFamiglia = NComFamiglia + RsDAtiTecComune("NRuoliSecondaCasa")
                End If
                
                SupMedieCalc(i + 1) = rs("Valore") * SuperficieMedia
                TotSupMedieCalc = TotSupMedieCalc + (SupMedieCalc(i + 1) * NComFamiglia)
                NComFamiglie(i + 1) = NComFamiglia
                rs.MoveNext
            Loop
            
        End If
        
        Alfa = TotSupMedieCalc / RsDAtiTecComune("TotSuperficieUtenzaDomestica")
        
        For i = 0 To UBound(SupMedieCalc)
            SupMedieCalc(i) = (SupMedieCalc(i) / Alfa) * NComFamiglie(i)
        Next
    Else
        Do While Not RsDAtiUD.EOF
                i = -1
                On Error Resume Next
                i = UBound(SupMedieCalc)
                On Error GoTo 0
                ReDim Preserve SupMedieCalc(i + 1)
            
            SupMedieCalc(i + 1) = RsDAtiUD("Superficie")

            RsDAtiUD.MoveNext
        Loop

    End If

    SupCorrettaSenzaRiduzione = SupMedieCalc
    
    If ConRiduzione Then
        For i = 0 To UBound(SupMedieCalc)
            TotaliDaRiduzioniUD i + 1, rsRIDUD, , TotSuperficieInRiduzioniDAUD_PF
            SupMedieCalc(i) = SupMedieCalc(i) - TotSuperficieInRiduzioniDAUD_PF
        Next
    End If

    SuperficeCorrettaUD = True
End Function

Function SelezionaCoeff507(TpoCoeff507 As TpoCoeff, ByVal CodComune, ByVal AnnoRiferimento As String) As ADODB.Recordset
    Dim qry As Command
    
    'Restituisce il coeff IPS e IQS delle Categorie UA
    Select Case TpoCoeff507
        Case CoeffUA507
            Set qry = gADOX.Procedures("SelCoeffIQSCategorieAttivita").Command
            qry.Parameters.Refresh
        
        Case CoeffUD507
            Set qry = gADOX.Procedures("SelCoeff507UtenzeDomestiche").Command
            qry.Parameters.Refresh
    End Select
    
    
    qry.Parameters("[pcodComune]") = CodComune
    qry.Parameters("[pAnnoRiferimento]") = AnnoRiferimento


    Set SelezionaCoeff507 = CreateObject("ADODB.Recordset")
    SelezionaCoeff507.CursorLocation = adUseClient
    SelezionaCoeff507.LockType = adLockBatchOptimistic
    SelezionaCoeff507.CursorType = adOpenStatic
    Set SelezionaCoeff507.ActiveConnection = gDBC
    
    Set SelezionaCoeff507.Source = qry
    SelezionaCoeff507.Open
    
    
    Set qry.ActiveConnection = Nothing
    Set SelezionaCoeff507.ActiveConnection = Nothing
End Function

Function SelezionaTotSvuotamenti(ByVal CodComune, ByVal AnnoRiferimento As String) As ADODB.Recordset
    Dim qry As Command
    
    'Restituisce il totale svuotamento dato il comune e l'anno di riferimento
    Set qry = gADOX.Procedures("SelTotSvuotamentiPerComune").Command
    qry.Parameters.Refresh
    
    
    qry.Parameters("pcodComune") = CodComune
    qry.Parameters("pAnnoRiferimento") = AnnoRiferimento

    Set SelezionaTotSvuotamenti = CreateObject("ADODB.Recordset")
    SelezionaTotSvuotamenti.CursorLocation = adUseClient
    SelezionaTotSvuotamenti.LockType = adLockBatchOptimistic
    SelezionaTotSvuotamenti.CursorType = adOpenStatic
    Set SelezionaTotSvuotamenti.ActiveConnection = gDBC
    
    Set SelezionaTotSvuotamenti.Source = qry
    SelezionaTotSvuotamenti.Open
    
    
    Set qry.ActiveConnection = Nothing
    Set SelezionaTotSvuotamenti.ActiveConnection = Nothing
End Function

Function SelezionaTotSvuotamentiUD(ByVal CodComune, ByVal AnnoRiferimento As String) As ADODB.Recordset
    Dim qry As Command
    
    'Restituisce il totale svuotamento UD dato il comune e l'anno di riferimento
    Set qry = gADOX.Procedures("SelTotSvuotamentiUDPerComune").Command
    qry.Parameters.Refresh
        
    
    
    qry.Parameters("pcodComune") = CodComune
    qry.Parameters("pAnnoRiferimento") = AnnoRiferimento


    Set SelezionaTotSvuotamentiUD = CreateObject("ADODB.Recordset")
    SelezionaTotSvuotamentiUD.CursorLocation = adUseClient
    SelezionaTotSvuotamentiUD.LockType = adLockBatchOptimistic
    SelezionaTotSvuotamentiUD.CursorType = adOpenStatic
    Set SelezionaTotSvuotamentiUD.ActiveConnection = gDBC
    
    Set SelezionaTotSvuotamentiUD.Source = qry
    SelezionaTotSvuotamentiUD.Open
    
    
    Set qry.ActiveConnection = Nothing
    Set SelezionaTotSvuotamentiUD.ActiveConnection = Nothing
End Function

Function SelezionaTotSvuotamentiUA(ByVal CodComune, ByVal AnnoRiferimento As String) As ADODB.Recordset
    Dim qry As Command
    
    'Restituisce il totale svuotamento UA dato il comune e l'anno di riferimento
    Set qry = gADOX.Procedures("SelTotSvuotamentiUAPerComune").Command
    qry.Parameters.Refresh
        
    
    
    qry.Parameters("pcodComune") = CodComune
    qry.Parameters("pAnnoRiferimento") = AnnoRiferimento


    Set SelezionaTotSvuotamentiUA = CreateObject("ADODB.Recordset")
    SelezionaTotSvuotamentiUA.CursorLocation = adUseClient
    SelezionaTotSvuotamentiUA.LockType = adLockBatchOptimistic
    SelezionaTotSvuotamentiUA.CursorType = adOpenStatic
    Set SelezionaTotSvuotamentiUA.ActiveConnection = gDBC
    
    Set SelezionaTotSvuotamentiUA.Source = qry
    SelezionaTotSvuotamentiUA.Open
    
    
    Set qry.ActiveConnection = Nothing
    Set SelezionaTotSvuotamentiUA.ActiveConnection = Nothing
End Function


Function SelezionaCoeffK(TpoCoeff As TpoCoeff, ByVal CodComune As String, ByVal AnnoRiferimento As String, ByVal AreaGeografica As String, ByVal ParAbitanti As String) As Recordset
    Dim qry As Command
    
    If TpoCoeff = CoeffKB Or TpoCoeff = CoeffKA Then
        'KA e KB
        Set qry = gADOX.Procedures("SelCoeffUtenzeDomesticheComunali").Command
    Else
        'KC e KD
        Set qry = gADOX.Procedures("SelCoeffCategorieAttivitaComunali").Command
    End If
    
    qry.Parameters.Refresh
    qry.Parameters("Comune") = CodComune
    qry.Parameters("[Anno]") = AnnoRiferimento
    qry.Parameters("[CodParGeo]") = AreaGeografica
    qry.Parameters("[CodR]") = ParAbitanti
    
    If TpoCoeff = CoeffKB Or TpoCoeff = CoeffKD Then
        'KB e KD
        qry.Parameters("ParteF") = False
    Else
        'KA e KC
        qry.Parameters("ParteF") = True
    End If
    
    Set SelezionaCoeffK = qry.Execute
End Function


Function SelezionaUtenzeDomestiche(ByVal CodComune As String, ByVal AnnoRiferimento As String) As Recordset
    Dim qry As Command
    
    Set qry = gADOX.Procedures("SelUtenzeDomesticheComunali").Command
    qry.Parameters.Refresh
    qry.Parameters("[pCodComune]") = CodComune
    qry.Parameters("[pAnno]") = AnnoRiferimento


    Set SelezionaUtenzeDomestiche = CreateObject("ADODB.Recordset")
    SelezionaUtenzeDomestiche.CursorLocation = adUseClient
    SelezionaUtenzeDomestiche.LockType = adLockBatchOptimistic
    SelezionaUtenzeDomestiche.CursorType = adOpenStatic
    Set SelezionaUtenzeDomestiche.ActiveConnection = gDBC
    
    Set SelezionaUtenzeDomestiche.Source = qry
    SelezionaUtenzeDomestiche.Open
    
    
    Set qry.ActiveConnection = Nothing
    Set SelezionaUtenzeDomestiche.ActiveConnection = Nothing

End Function


Function SelezionaDatiFinali507UA(ByVal CodComune As String, ByVal AnnoRiferimento As String) As Recordset
    Dim qry As Command
    
    Set qry = gADOX.Procedures("SelDatiFinali507UA").Command
    qry.Parameters.Refresh
    qry.Parameters("[pCodComune]") = CodComune
    qry.Parameters("[pAnno]") = AnnoRiferimento

    Set SelezionaDatiFinali507UA = qry.Execute

End Function

Function SelezionaDatiFinali507UD(ByVal CodComune As String, ByVal AnnoRiferimento As String) As Recordset
    Dim qry As Command
    
    Set qry = gADOX.Procedures("SelDatiFinali507UD").Command
    qry.Parameters.Refresh
    qry.Parameters("[pCodComune]") = CodComune
    qry.Parameters("[pAnno]") = AnnoRiferimento

    Set SelezionaDatiFinali507UD = qry.Execute

End Function


Function SelezionaUtenzeAttivita(ByVal CodComune As String, ByVal AnnoRiferimento As String) As Recordset
    Dim qry As Command
    
    Set qry = gADOX.Procedures("SelCategorieAttivitaComunali").Command
    qry.Parameters.Refresh
    qry.Parameters("[pCodComune]") = CodComune
    qry.Parameters("[pAnno]") = AnnoRiferimento

    Set SelezionaUtenzeAttivita = qry.Execute

End Function


Function SelezionaBonus(ByVal Tipo As TpoBonus, ByVal CodComune As String, ByVal AnnoRiferimento As String, ByVal pCodTipoTariffa As tpoTariffa) As Recordset
    Dim qry As Command
    
    If Tipo = BonusUtenzaAttivita Then
        Set qry = gADOX.Procedures("SelBonusCategorieAttivita").Command
    Else
        Set qry = gADOX.Procedures("SelBonusUtenzeDomestiche").Command
    End If
        
    qry.Parameters.Refresh
    qry.Parameters("pCodComune") = CodComune
    qry.Parameters("pAnnoRiferimento") = AnnoRiferimento
    qry.Parameters("pCodTipoTariffa") = pCodTipoTariffa

    Set SelezionaBonus = qry.Execute

End Function

Function SelezionaContenitoriSvuotamentiIncarico(ByVal Tipo As TpoContenitoriSvuotaInDotazione, ByVal CodComune As String, ByVal AnnoRiferimento As String) As Recordset
    Dim qry As ADODB.Command
    Dim p  As ADODB.Parameter
    
    Set qry = CreateObject("ADODB.Command")
    Select Case Tipo
        Case TpoContenitoriSvuotaInDotazione.ContenitoriInDotazioneUA
            qry.CommandText = "select ConDot.*,CA.Descrizione,CONT.Descrizione AS DesCont from ((ContenitoriInDotazioneCAComunali as ConDot  " & _
                " inner join CategorieAttivitaComunali as CA on CA.codComune = ConDot.codComune and ConDot.annoriferimento = CA.annoriferimento" & _
                " and ConDot.CodAttivita = CA.idCategoriaAttivita)" & _
                " left join Contenitori AS CONT on CONT.id = ConDot.idContenitore) " & _
                " where ConDot.annoriferimento=? and ConDot.codcomune=? and ConDot.flgSimulazione=false"
        Case TpoContenitoriSvuotaInDotazione.ContenitoriInDotazioneUD
            qry.CommandText = "select ConDot.*,NC.Descrizione,CONT.Descrizione AS DesCont from ((ContenitoriInDotazioneUDComunali as ConDot " & _
                " inner join NumeroComponentiFamiglia as NC on ConDot.IdNumeriComponenti = NC.NumeroComponenti)" & _
                " left join Contenitori AS CONT on CONT.id = ConDot.idContenitore) " & _
                " where ConDot.annoriferimento=? and ConDot.codcomune=? and ConDot.flgSimulazione=false"
        Case TpoContenitoriSvuotaInDotazione.ContenitoriSingoliSvuotamentiUA
            qry.CommandText = "select ConSvu.*,CA.Descrizione,CONT.Descrizione AS DesCont  from ((ContenitoriSvuotamentiCAComunali AS ConSvu " & _
                " inner join CategorieAttivitaComunali as CA on CA.codComune = ConSvu.codComune and ConSvu.annoriferimento = CA.annoriferimento" & _
                " and ConSvu.CodAttivita = CA.idCategoriaAttivita)" & _
                " left join Contenitori AS CONT on CONT.id = ConSvu.idContenitore) " & _
                " where ConSvu.annoriferimento=? and ConSvu.codcomune=? and ConSvu.flgSimulazione=false"
        Case TpoContenitoriSvuotaInDotazione.ContenitoriSingoliSvuotamentiUD
            qry.CommandText = "select ConSvu.*,NC.Descrizione,CONT.Descrizione AS DesCont from ((ContenitoriSvuotamentiUDComunali AS ConSvu " & _
                " inner join NumeroComponentiFamiglia as NC on ConSvu.IdNumeriComponenti = NC.NumeroComponenti)" & _
                " left  join Contenitori AS CONT on CONT.id = ConSvu.idContenitore) " & _
                " where ConSvu.annoriferimento=? and ConSvu.codcomune=? and ConSvu.flgSimulazione=false"
        
        Case TpoContenitoriSvuotaInDotazione.ContenitoriSvuotamentiUA
            qry.CommandText = "select * from ContenitoriSvuotamentiCAComunali where annoriferimento=? and codcomune=? and IdContenitore=0"
        Case TpoContenitoriSvuotaInDotazione.ContenitoriSvuotamentiUD
            qry.CommandText = "select * from ContenitoriSvuotamentiUDComunali where annoriferimento=? and codcomune=? and IdContenitore=0"
    End Select
        
    qry.CommandType = adCmdText
    Set qry.ActiveConnection = gDBC
    qry.Parameters.Refresh
    
    qry.Parameters(0) = AnnoRiferimento
    qry.Parameters(1) = CodComune


    Set SelezionaContenitoriSvuotamentiIncarico = qry.Execute

End Function



Function InValutaEst(ByVal Valore As String, valuta As String, Optional numDec As Integer = 2) As String
    Dim FORMATVALUTA As String
    
    If Trim$(Valore) = vbNullString Then Exit Function
    
    FORMATVALUTA = "###,###,###,###,##0." & String(numDec, "0")
    Select Case valuta
        Case VALUTA_LIRE
            InValutaEst = Format$(InLire(CDbl(Valore)), FORMATSTANDARD)
        Case VALUTA_EURO
            InValutaEst = Format$(InEuro(CDbl(Valore)), FORMATVALUTA)
        Case Else
            If gValutaDefault = VALUTA_EURO Then
                InValutaEst = Format$(CDbl(Valore), FORMATVALUTA)
            Else
                InValutaEst = Format$(CDbl(Valore), FORMATSTANDARD)
            End If
    End Select
End Function


Sub TotaliDaRiduzioniUA(ChiaveUA As String, rsRIDUA As ADODB.Recordset, Optional TotSuperficieInRiduzioniDAUA_Pv As Double, Optional TotSuperficieInRiduzioniDAUA_PF As Double)
    
    If Not rsRIDUA Is Nothing Then rsRIDUA.Filter = ""
    If rsRIDUA.EOF Then Exit Sub
    
    TotSuperficieInRiduzioniDAUA_Pv = 0
    TotSuperficieInRiduzioniDAUA_PF = 0
    
    rsRIDUA.MoveFirst
    
    If ChiaveUA <> "" Then
        rsRIDUA.Filter = "CodAttivita=" & ChiaveUA
    End If
    
    
    If Not rsRIDUA.EOF Then
        Do While Not rsRIDUA.EOF
            If rsRIDUA("ParteFissa") = True Then
                TotSuperficieInRiduzioniDAUA_PF = TotSuperficieInRiduzioniDAUA_PF + rsRIDUA("Valore") * (rsRIDUA("Riduzione") / 100)
            Else
                TotSuperficieInRiduzioniDAUA_Pv = TotSuperficieInRiduzioniDAUA_Pv + rsRIDUA("Valore") * (rsRIDUA("Riduzione") / 100)
            End If
            rsRIDUA.MoveNext
        Loop
        rsRIDUA.MoveFirst
    End If
    
    rsRIDUA.Filter = ""
End Sub


Sub TotaliDaRiduzioniUD(ByVal ChiaveUD As String, ByVal rsRIDUD As ADODB.Recordset, Optional TotUtenzeInRiduzioniDAUD_PV As Double, Optional TotSuperficieInRiduzioniDAUD_PF As Double)
    Dim i As Integer
    
    If Not rsRIDUD Is Nothing Then rsRIDUD.Filter = ""
    If rsRIDUD.EOF Then Exit Sub
    TotUtenzeInRiduzioniDAUD_PV = 0
    TotSuperficieInRiduzioniDAUD_PF = 0
    
    
    rsRIDUD.MoveFirst
    
    If ChiaveUD <> "" Then
        rsRIDUD.Filter = "IdNumeriComponenti= " & ChiaveUD
    End If
    
    
    If Not rsRIDUD.EOF Then
        Do While Not rsRIDUD.EOF
            If rsRIDUD("ParteFissa") = True Then
                TotSuperficieInRiduzioniDAUD_PF = TotSuperficieInRiduzioniDAUD_PF + rsRIDUD("Valore") * (rsRIDUD("Riduzione") / 100)
            Else
                TotUtenzeInRiduzioniDAUD_PV = TotUtenzeInRiduzioniDAUD_PV + rsRIDUD("Valore") * (rsRIDUD("Riduzione") / 100)
            End If
            rsRIDUD.MoveNext
        Loop
        rsRIDUD.MoveFirst
    End If
    
    rsRIDUD.Filter = ""

End Sub



