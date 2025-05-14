Attribute VB_Name = "ModSimulazioni"
Option Explicit

Type Simulation
    ID As String
    description As String
End Type

'Creazione di una simulazione - per poter creare una simulazione di un anno tariffario
'è necessario calcolare la tariffa, una volta fatto la creazione sarà data dalla copia
'dei seguenti archivi , generando un codice anno + _ + un progressivo simulazione
'(il numero massimo di simulazioni e 9).
Sub createSimulation(idAnno As String, CodComune As String, TipoMetodo As Integer)
    On Error GoTo createSimulationERR:
    Dim commit As Boolean
    Dim mSimulazione As Simulation
    gDBC.BeginTrans
    commit = True
    
    mSimulazione = getSimulazione(CodComune, idAnno)
    If mSimulazione.ID = "-1" Then Exit Sub
    
    mSimulazione.description = Replace(mSimulazione.description, "'", "''")
    'duplicazione archivi
    'Anagrafica Comunale
    gDBC.Execute ("INSERT INTO AnagraficaComunale ( CodComune, AnnoRiferimento, NumAbitantiResidenti, NumNucleoFamiliare, FamigliaCon1Componente, " & _
        "FamigliaCon2Componenti, FamigliaCon3Componenti, FamigliaCon4Componenti, FamigliaCon5Componenti, " & _
        "FamigliaConPiuDi5Componenti, Immigrati, Emmigrati, NumMedioResidenti, NumDomiciliatii, NumMedioTuristi, " & _
        "SuperficieDelTerrComunale, AmbitoComunaleOttimale, PRComunaleOttimale,flgSimulazione) " & _
        "SELECT anaCom.CodComune,'" & mSimulazione.ID & "', anaCom.NumAbitantiResidenti, anaCom.NumNucleoFamiliare, " & _
        "anaCom.FamigliaCon1Componente , anaCom.FamigliaCon2Componenti, anaCom.FamigliaCon3Componenti, anaCom.FamigliaCon4Componenti, " & _
        "anaCom.FamigliaCon5Componenti , anaCom.FamigliaConPiuDi5Componenti, anaCom.Immigrati, anaCom.Emmigrati, anaCom.NumMedioResidenti, " & _
        "anaCom.NumDomiciliatii , anaCom.NumMedioTuristi, anaCom.SuperficieDelTerrComunale, anaCom.AmbitoComunaleOttimale, " & _
        "anaCom.PRComunaleOttimale,true FROM AnagraficaComunale as anaCom where anaCom.CodComune='" & CodComune & "' and anaCom.AnnoRiferimento='" & idAnno & "'")
    'DatiTecniciComune
    gDBC.Execute ("INSERT INTO DatiTecniciComune ( CodComune, QRifiutiTotali, QRifiutiRaccoltaDifferenziata, QRifiutiCategorieAttivita, NRuoliUtenzaDomestica, NRuoliUtenzaAttivita, TotSuperficieUtenzaDomestica, TotSuperficieUtenzaAttivita, TotSuperficieNonIscritti, TotGettitoDaUtenzaDomestica, TotGettitoDaUtenzaAttivita, NRuoliSecondaCasa, Id_NumCompSecondaCasa,  AnnoRiferimento, flgSimulazione ) " & _
    "SELECT DatiTecniciComune.CodComune, DatiTecniciComune.QRifiutiTotali, DatiTecniciComune.QRifiutiRaccoltaDifferenziata, DatiTecniciComune.QRifiutiCategorieAttivita, " & _
    "DatiTecniciComune.NRuoliUtenzaDomestica, DatiTecniciComune.NRuoliUtenzaAttivita, DatiTecniciComune.TotSuperficieUtenzaDomestica, DatiTecniciComune.TotSuperficieUtenzaAttivita, DatiTecniciComune.TotSuperficieNonIscritti, DatiTecniciComune.TotGettitoDaUtenzaDomestica, DatiTecniciComune.TotGettitoDaUtenzaAttivita, DatiTecniciComune.NRuoliSecondaCasa, DatiTecniciComune.Id_NumCompSecondaCasa,'" & mSimulazione.ID & "', True " & _
    "From DatiTecniciComune " & _
    "WHERE (((DatiTecniciComune.CodComune)='" & CodComune & "') AND ((DatiTecniciComune.AnnoRiferimento)='" & idAnno & "'))")
    'CostoComunale
    gDBC.Execute ("INSERT INTO CostoComunale ( CodComune, AnnoRiferimento, Tot_Servizio, Gettito, CG, CSL, CRT, CTS, AC, CGIND, CRD, CTR_COSTI, CTR_RICAVI, CTR, CON, Y, CGD, CC, CARC, CGG, CCD, CK, AMM, ACC, R, Rn, Kn_Anno_Prec, Investimenti_Programmati, Fn, Tot_CSL, Tot_CRT, Tot_CRD, Tot_CTS, Tot_CTR, flgSimulazione ) " & _
    "SELECT CostoComunale.CodComune, '" & mSimulazione.ID & "', CostoComunale.Tot_Servizio, CostoComunale.Gettito, CostoComunale.CG, CostoComunale.CSL, " & _
    "CostoComunale.CRT, CostoComunale.CTS, CostoComunale.AC, CostoComunale.CGIND, CostoComunale.CRD, CostoComunale.CTR_COSTI, CostoComunale.CTR_RICAVI, " & _
    "CostoComunale.CTR, CostoComunale.CON, CostoComunale.Y, CostoComunale.CGD, CostoComunale.CC, CostoComunale.CARC, CostoComunale.CGG, " & _
    "CostoComunale.CCD, CostoComunale.CK, CostoComunale.AMM, CostoComunale.ACC, CostoComunale.R, CostoComunale.Rn, CostoComunale.Kn_Anno_Prec, " & _
    "CostoComunale.Investimenti_Programmati, CostoComunale.Fn, CostoComunale.Tot_CSL, CostoComunale.Tot_CRT, CostoComunale.Tot_CRD, CostoComunale.Tot_CTS, CostoComunale.Tot_CTR, True " & _
    "From CostoComunale " & _
    "WHERE (((CostoComunale.CodComune)='" & CodComune & "') AND ((CostoComunale.AnnoRiferimento)='" & idAnno & "'))")
    
    
    'crea record in simulazioni
    gDBC.Execute ("insert into Simulazioni (codComune,idSimulazione,descrizione,TipoMetodo) values ('" & CodComune & "' , '" & mSimulazione.ID & "' , '" & mSimulazione.description & "'," & TipoMetodo & ")")
    
    'Accantonamenti
    gDBC.Execute ("INSERT INTO Accantonamenti ( CodComune, AnnoRiferimento, CodAccantonamento, CostoUnitario, Quantita, Totale, flgSimulazione ) " & _
    "SELECT Accantonamenti.CodComune, '" & mSimulazione.ID & "', Accantonamenti.CodAccantonamento, Accantonamenti.CostoUnitario, Accantonamenti.Quantita, Accantonamenti.Totale, true " & _
    "From Accantonamenti " & _
    "WHERE (((Accantonamenti.CodComune)='" & CodComune & "') AND ((Accantonamenti.AnnoRiferimento)='" & idAnno & "'))")
    'Ammortamenti
    gDBC.Execute ("INSERT INTO Ammortamenti ( CodComune,  CodTipoAmmortamento, Quantita, ValoreCommerciale, AnniAmmortamento, PercValoreResiduo, PercInteresseComputo, QuotaAnnua, flgSimulazione, AnnoRiferimento ) " & _
    "SELECT Ammortamenti.CodComune,  Ammortamenti.CodTipoAmmortamento, Ammortamenti.Quantita, Ammortamenti.ValoreCommerciale, " & _
    "Ammortamenti.AnniAmmortamento, Ammortamenti.PercValoreResiduo, Ammortamenti.PercInteresseComputo, Ammortamenti.QuotaAnnua, True , '" & mSimulazione.ID & "'" & _
    " From Ammortamenti WHERE (((Ammortamenti.CodComune)='" & CodComune & "') AND ((Ammortamenti.AnnoRiferimento)='" & idAnno & "'))")
    'PersonaleImpiegatoNelComune
    gDBC.Execute ("INSERT INTO PersonaleImpiegatoNelComune ( CodComune, CodTipoPersonale, Liv2, Liv3, Liv4, Liv5, Liv6, Liv7, Liv7q, Liv8, Liv8q, Amministrativo, Manager, AnnoRiferimento, flgSimulazione ) " & _
    "SELECT PersonaleImpiegatoNelComune.CodComune, PersonaleImpiegatoNelComune.CodTipoPersonale, PersonaleImpiegatoNelComune.Liv2, " & _
    "PersonaleImpiegatoNelComune.Liv3, PersonaleImpiegatoNelComune.Liv4, PersonaleImpiegatoNelComune.Liv5, PersonaleImpiegatoNelComune.Liv6, PersonaleImpiegatoNelComune.Liv7, PersonaleImpiegatoNelComune.Liv7q, PersonaleImpiegatoNelComune.Liv8, PersonaleImpiegatoNelComune.Liv8q, PersonaleImpiegatoNelComune.Amministrativo, PersonaleImpiegatoNelComune.Manager,'" & mSimulazione.ID & "', True " & _
    "From PersonaleImpiegatoNelComune " & _
    "WHERE (((PersonaleImpiegatoNelComune.CodComune)='" & CodComune & "') AND ((PersonaleImpiegatoNelComune.AnnoRiferimento)='" & idAnno & "'))")
    'RetribuzioneDelPersonale
    gDBC.Execute ("INSERT INTO RetribuzioneDelPersonale ( CodComune, CodTipoPersonale, CodLivelloPersonale, RetribuzioneAnnuale, AnnoRiferimento, flgSimulazione ) " & _
    "SELECT RetribuzioneDelPersonale.CodComune, RetribuzioneDelPersonale.CodTipoPersonale, RetribuzioneDelPersonale.CodLivelloPersonale, RetribuzioneDelPersonale.RetribuzioneAnnuale,'" & mSimulazione.ID & "', True " & _
    "From RetribuzioneDelPersonale " & _
    "WHERE (((RetribuzioneDelPersonale.CodComune)='" & CodComune & "') AND ((RetribuzioneDelPersonale.AnnoRiferimento)='" & idAnno & "'))")
    
    
    'BonusUtenzeAttivita
    gDBC.Execute ("INSERT INTO BonusUtenzeAttivita ( CodComune, CodAttivita, AnnoRiferimento, Riduzione, ParteFissa, Valore, CodTipoTariffa, Descrizione, flgSimulazione ) " & _
    "SELECT BonusUtenzeAttivita.CodComune, BonusUtenzeAttivita.CodAttivita, '" & mSimulazione.ID & "', BonusUtenzeAttivita.Riduzione, BonusUtenzeAttivita.ParteFissa, BonusUtenzeAttivita.Valore, " & _
    "BonusUtenzeAttivita.CodTipoTariffa, BonusUtenzeAttivita.Descrizione, True " & _
    "From BonusUtenzeAttivita " & _
    "WHERE BonusUtenzeAttivita.CodComune='" & CodComune & "' AND BonusUtenzeAttivita.AnnoRiferimento='" & idAnno & "'")
    'BonusUtenzeDomestiche
    gDBC.Execute ("INSERT INTO BonusUtenzeDomestiche ( AnnoRiferimento, flgSimulazione, CodComune, IdNumeriComponenti, Riduzione, ParteFissa, Valore, CodTipoTariffa, Descrizione, indSingle) " & _
    "SELECT '" & mSimulazione.ID & "', True, BonusUtenzeDomestiche.CodComune, BonusUtenzeDomestiche.IdNumeriComponenti, BonusUtenzeDomestiche.Riduzione, BonusUtenzeDomestiche.ParteFissa, BonusUtenzeDomestiche.Valore, BonusUtenzeDomestiche.CodTipoTariffa, BonusUtenzeDomestiche.Descrizione, BonusUtenzeDomestiche.indSingle " & _
    "From BonusUtenzeDomestiche " & _
    "WHERE BonusUtenzeDomestiche.CodComune='" & CodComune & "' AND BonusUtenzeDomestiche.AnnoRiferimento='" & idAnno & "'")
    
    'CategorieAttivitaComunali
    gDBC.Execute ("INSERT INTO CategorieAttivitaComunali ( IdCategoriaAttivita, Descrizione, CodComune, AnnoRiferimento, Superficie, NumRuoli, Tariffa, TariffaPF, TariffaPV, SuperficieM, flgSimulazione ) " & _
    "SELECT CategorieAttivitaComunali.IdCategoriaAttivita, CategorieAttivitaComunali.Descrizione, CategorieAttivitaComunali.CodComune, " & _
    "'" & mSimulazione.ID & "', CategorieAttivitaComunali.Superficie, CategorieAttivitaComunali.NumRuoli, CategorieAttivitaComunali.Tariffa, " & _
    "CategorieAttivitaComunali.TariffaPF, CategorieAttivitaComunali.TariffaPV, CategorieAttivitaComunali.SuperficieM, true " & _
    "From CategorieAttivitaComunali " & _
    "WHERE (((CategorieAttivitaComunali.CodComune)='" & CodComune & "') AND ((CategorieAttivitaComunali.AnnoRiferimento)='" & idAnno & "'))")
    'ParametriSistema
    gDBC.Execute ("INSERT INTO ParametriSistema ( CodIstat, Parametro, Descrizione, Valore, ValoreMin, ValoreMax, Privata, AnnoRiferimento, flgSimulazione ) " & _
    "SELECT ParametriSistema.CodIstat, ParametriSistema.Parametro, ParametriSistema.Descrizione, ParametriSistema.Valore, ParametriSistema.ValoreMin, ParametriSistema.ValoreMax, ParametriSistema.Privata, '" & mSimulazione.ID & "', True " & _
    "From ParametriSistema " & _
    "WHERE (((ParametriSistema.CodIstat)='" & CodComune & "') AND ((ParametriSistema.AnnoRiferimento)='" & idAnno & "'))")
    'UtenzeDomesticheComunali
    gDBC.Execute ("INSERT INTO UtenzeDomesticheComunali ( CodComune, IdNumeriComponenti, TariffaNComp, TariffaSuperficie, Superficie,NumeroUtenze, SuperficieM, AnnoRiferimento, flgSimulazione ) " & _
    "SELECT UtenzeDomesticheComunali.CodComune, UtenzeDomesticheComunali.IdNumeriComponenti, UtenzeDomesticheComunali.TariffaNComp, UtenzeDomesticheComunali.TariffaSuperficie, UtenzeDomesticheComunali.Superficie,NumeroUtenze, UtenzeDomesticheComunali.SuperficieM, '" & mSimulazione.ID & "', True " & _
    "From UtenzeDomesticheComunali " & _
    "WHERE (((UtenzeDomesticheComunali.CodComune)='" & CodComune & "') AND ((UtenzeDomesticheComunali.AnnoRiferimento)='" & idAnno & "'))")
    'ValoriCoefficienti507UDComunali
    gDBC.Execute ("INSERT INTO ValoriCoefficienti507UDComunali ( CodComune, CodUtenza, Valore,ValoreCalc, AnnoRiferimento, flgSimulazione ) " & _
    "SELECT ValCoeff507UD.CodComune, ValCoeff507UD.CodUtenza, ValCoeff507UD.Valore,ValCoeff507UD.ValoreCalc, '" & mSimulazione.ID & "', True " & _
    "From ValoriCoefficienti507UDComunali as ValCoeff507UD " & _
    "WHERE (((ValCoeff507UD.CodComune)='" & CodComune & "') AND ((ValCoeff507UD.AnnoRiferimento)='" & idAnno & "'))")
    'ValoriCoefficientiCAComunali
    gDBC.Execute ("INSERT INTO ValoriCoefficientiCAComunali ( IdCoefficiente, CodComune, Valore, AnnoRiferimento, flgSimulazione ) " & _
    "SELECT ValoriCoefficientiCAComunali.IdCoefficiente, ValoriCoefficientiCAComunali.CodComune, ValoriCoefficientiCAComunali.Valore, '" & mSimulazione.ID & "', True " & _
    "From ValoriCoefficientiCAComunali " & _
    "WHERE (((ValoriCoefficientiCAComunali.CodComune)='" & CodComune & "') AND ((ValoriCoefficientiCAComunali.AnnoRiferimento)='" & idAnno & "'))")
    'ValoriCoefficientiIQSCAComunali
    gDBC.Execute ("INSERT INTO ValoriCoefficientiIQSCAComunali ( CodComune, CodAttivita, IQS, IPS,IPSCalc, AnnoRiferimento, flgSimulazione ) " & _
    "SELECT ValCoeff507CA.CodComune, ValCoeff507CA.CodAttivita, ValCoeff507CA.IQS, ValCoeff507CA.IPS,ValCoeff507CA.IPSCalc, '" & mSimulazione.ID & "', True " & _
    "From ValoriCoefficientiIQSCAComunali as ValCoeff507CA " & _
    "WHERE (((ValCoeff507CA.CodComune)='" & CodComune & "') AND ((ValCoeff507CA.AnnoRiferimento)='" & idAnno & "'))")
    'ValoriCoefficientiUDComunali
    gDBC.Execute ("INSERT INTO ValoriCoefficientiUDComunali ( CodComune, IdCoefficiente, Valore, AnnoRiferimento, flgSimulazione ) " & _
    "SELECT ValoriCoefficientiUDComunali.CodComune, ValoriCoefficientiUDComunali.IdCoefficiente, ValoriCoefficientiUDComunali.Valore,'" & mSimulazione.ID & "', True " & _
    "From ValoriCoefficientiUDComunali " & _
    "WHERE (((ValoriCoefficientiUDComunali.CodComune)='" & CodComune & "') AND ((ValoriCoefficientiUDComunali.AnnoRiferimento)='" & idAnno & "'))")
    
    'DatiFinaliCalcoloTariffa
    gDBC.Execute ("INSERT INTO DatiFinaliCalcoloTariffa ( CodComune,CodTipoTariffa, AnnoRiferimento, QVDOM, QFDOM, QUVDOM, QUFDOM, CUDOM, QVNDOM, QFNDOM, CUNDOM, QAPFNDOM,GettitoUD,GettitoUA, IntroitoDaUA_PF,IntroitoDaUA_PV,IntroitoDaUD_PV,IntroitoDaUD_PF, flgSimulazione ) " & _
    "SELECT DatiFinaliCalcoloTariffa.CodComune,DatiFinaliCalcoloTariffa.CodTipoTariffa, '" & mSimulazione.ID & "', DatiFinaliCalcoloTariffa.QVDOM, DatiFinaliCalcoloTariffa.QFDOM, DatiFinaliCalcoloTariffa.QUVDOM, DatiFinaliCalcoloTariffa.QUFDOM, DatiFinaliCalcoloTariffa.CUDOM, DatiFinaliCalcoloTariffa.QVNDOM, " & _
    "DatiFinaliCalcoloTariffa.QFNDOM, DatiFinaliCalcoloTariffa.CUNDOM, DatiFinaliCalcoloTariffa.QAPFNDOM,DatiFinaliCalcoloTariffa.GettitoUD,DatiFinaliCalcoloTariffa.GettitoUA, DatiFinaliCalcoloTariffa.IntroitoDaUA_PF,DatiFinaliCalcoloTariffa.IntroitoDaUA_PV,DatiFinaliCalcoloTariffa.IntroitoDaUD_PV,DatiFinaliCalcoloTariffa.IntroitoDaUD_PF, True AS Espr2 " & _
    "From DatiFinaliCalcoloTariffa " & _
    "WHERE DatiFinaliCalcoloTariffa.CodComune='" & CodComune & "' AND DatiFinaliCalcoloTariffa.AnnoRiferimento='" & idAnno & "'")
    'DatiFinaliCalcoloTarsu507
    gDBC.Execute ("INSERT INTO DatiFinaliCalcoloTarsu507 ( CodComune, CostoTotale, Gettito, CMG,CoeffCorrettivoIPS,IQSUtenzaDomestica,IPSUtenzaDomestica, ProdMedia, Copertura, AnnoRiferimento, flgSimulazione ) " & _
    "SELECT DatiFinaliCalcoloTarsu507.CodComune, DatiFinaliCalcoloTarsu507.CostoTotale, DatiFinaliCalcoloTarsu507.Gettito, " & _
    "DatiFinaliCalcoloTarsu507.CMG,DatiFinaliCalcoloTarsu507.CoeffCorrettivoIPS,IQSUtenzaDomestica,IPSUtenzaDomestica, DatiFinaliCalcoloTarsu507.ProdMedia, DatiFinaliCalcoloTarsu507.Copertura,'" & mSimulazione.ID & "', True " & _
    "From DatiFinaliCalcoloTarsu507 " & _
    "WHERE DatiFinaliCalcoloTarsu507.CodComune='" & CodComune & "' AND DatiFinaliCalcoloTarsu507.AnnoRiferimento='" & idAnno & "'")
    gDBC.CommitTrans
    Exit Sub
createSimulationERR:
    
    If commit Then gDBC.RollbackTrans
    
    MsgBox Error
    
End Sub

Function getSimulazione(CodComune As String, AnnoRiferimento As String) As Simulation
    'chiedo descrizione per la simulazione
    Dim idSimu As Integer
    Dim rs As ADODB.Recordset
    Dim s As String
    
    getSimulazione.ID = "-1"
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockBatchOptimistic
    Set rs.ActiveConnection = gDBC
    
    rs.Source = "select right(idSimulazione,1) as id from Simulazioni where idSimulazione like '" & AnnoRiferimento & "_%' and codComune = '" & CodComune & "' order by idsimulazione"
    rs.Open
    Set rs.ActiveConnection = Nothing
    
    If rs.EOF Then
        idSimu = 1
    Else
        If rs.RecordCount = 9 Then
            MsgBox "Attenzione superato il numero massimo di simulazioni (max=9)"
            Exit Function
        Else
            
            idSimu = rs(0)
            Do While Not rs.EOF
                
                If rs(0) > idSimu Then
                    Exit Do
                End If
                
                rs.MoveNext
                idSimu = idSimu + 1
            Loop
           
       End If
    End If
    
    rs.Close
    
    s = InputBox("Dammi la descrizione per la simulazione", "Simulazione", getDesComune(CodComune) & " " & Now)
    If Trim$(s) = "" Then s = getDesComune(CodComune) & " " & Now
    getSimulazione.description = s
    getSimulazione.ID = AnnoRiferimento & "_" & idSimu
End Function



Sub deleteSimulation(idSimulazione As String, CodComune As String, TipoMetodo As Integer)
    Dim bCommit As Boolean
    
    On Error GoTo deleteSimulationERR
    gDBC.BeginTrans
    bCommit = True
    
    gDBC.Execute ("delete from Simulazioni where CodComune ='" & CodComune & "' and idSimulazione='" & idSimulazione & "' and TipoMetodo=" & TipoMetodo)
    gDBC.Execute ("delete from  AnagraficaComunale where CodComune ='" & CodComune & "' and annoriferimento='" & idSimulazione & "'")
    gDBC.Execute ("delete from  CostoComunale where CodComune ='" & CodComune & "' and  annoriferimento='" & idSimulazione & "'")
    gDBC.Execute ("delete from  DatiTecniciComune where CodComune ='" & CodComune & "' and  annoriferimento='" & idSimulazione & "'")
    gDBC.Execute ("delete from  ParametriSistema where codIstat ='" & CodComune & "' and  annoriferimento='" & idSimulazione & "'")
    gDBC.CommitTrans
    Exit Sub
deleteSimulationERR:
    
    If bCommit Then gDBC.RollbackTrans
    MsgBox Error, vbCritical, App.Title
End Sub



Function replacePath(cmdtxt As String, spath As String) As String
Dim pos1 As Integer
Dim pos2 As Integer
Dim cmdnew As String
    
    pos1 = InStr(cmdtxt, "`")
    pos2 = InStr(pos1 + 1, cmdtxt, "`")
    
    Do While pos1 > 0
     '   Mid$(cmdtxt, pos1 + 1, pos2 - pos1) = spath
        
        cmdnew = cmdnew & Left(cmdtxt, pos1) & spath
        cmdtxt = Mid$(cmdtxt, pos2)

         
        pos1 = InStr(2, cmdtxt, "`")
        pos2 = InStr(pos1 + 1, cmdtxt, "`")
    Loop
    replacePath = cmdnew & cmdtxt
End Function

Sub ReportTariffa(idSimulazione As String, CodComune As String, TipoMetodo As tpoTariffa)
    Dim Anno As String
    Dim xlsApp As Object
    Dim wrkBook  As Object
    Dim qryTables As Object
    Dim i As Integer
    On Error GoTo ReportTariffaERR
    
    Anno = Left$(idSimulazione, Len(idSimulazione) - 2)
    
    Set xlsApp = CreateObject("Excel.Application")
    If TipoMetodo = Tariffa507 Then
        Set wrkBook = xlsApp.Workbooks.Open(App.Path & "\ReportSimulazioni507.xls")
    Else
        Set wrkBook = xlsApp.Workbooks.Open(App.Path & "\ReportSimulazioni.xls")
    End If
    
    
    
    xlsApp.DisplayAlerts = False
    Screen.MousePointer = vbHourglass
    wrkBook.Sheets("tariffa").Range("A1") = Anno
    wrkBook.Sheets("tariffa").Range("B1") = CodComune
    wrkBook.Sheets("tariffa").Range("F1") = TipoMetodo
    
    
    Dim stringaConnessione As String
    
    
    stringaConnessione = "ODBC;DSN=Simulazione;" & _
     "DBQ=" & App.Path & "\StudioAmbiente2002.mdb;MaxBufferSize=2048;PageTimeout=5;PWD=" & gPasswordDB & ";UID=admin;"
   
   
   
    'xlsApp.DisplayAlerts = True
    'xlsApp.Visible = True
    Set qryTables = wrkBook.Sheets("tariffa").QueryTables
   

    
    For i = 1 To qryTables.Count
       ' qryTables(i).SavePassword = True
       ' Debug.Print qryTables(i).Connection
       Debug.Print qryTables(i).Name
 

        qryTables(i).CommandText = replacePath(qryTables(i).CommandText, App.Path & "\StudioAmbiente2002")
        qryTables(i).Connection = stringaConnessione
        qryTables(i).Refresh False
        DoEvents
    Next
    
   
    
    Set qryTables = wrkBook.Sheets("simulazioneUD").QueryTables
    
    'qryTables("Simulazioni").SavePassword = True
    qryTables("Simulazioni").CommandText = replacePath(qryTables("Simulazioni").CommandText, App.Path & "\StudioAmbiente2002")
    qryTables("Simulazioni").Connection = stringaConnessione
    qryTables("Simulazioni").Refresh False
    
    
    For i = 1 To qryTables.Count
        'qryTables(i).SavePassword = True
        qryTables(i).CommandText = replacePath(qryTables(i).CommandText, App.Path & "\StudioAmbiente2002")
        qryTables(i).Connection = stringaConnessione
        qryTables(i).Refresh False
        DoEvents
    Next
    
    Set qryTables = wrkBook.Sheets("simulazioneCA").QueryTables
    For i = 1 To qryTables.Count
        'qryTables(i).SavePassword = True
        qryTables(i).CommandText = replacePath(qryTables(i).CommandText, App.Path & "\StudioAmbiente2002")
        qryTables(i).Connection = stringaConnessione
        qryTables(i).Refresh False
        DoEvents
    Next
  
    wrkBook.Application.CalculateFull
    
    Screen.MousePointer = vbDefault
    wrkBook.Application.Visible = True
    
    
    Exit Sub
ReportTariffaERR:

    Screen.MousePointer = vbDefault
    MsgBox Error, vbCritical, App.Title
    On Error Resume Next
    wrkBook.Application.Quit
End Sub

Sub PublicaComeTariffa(idSimulazione As String, CodComune As String, TipoMetodo As Integer)
    Dim bCommit As Boolean
    Dim Anno As String
    Dim rs As ADODB.Recordset
    
    On Error GoTo PublicaComeTariffaERR
    gDBC.BeginTrans
    bCommit = True
    
    'aggiorno la tariffa con  i dati della simulazione
    If MsgBox("Attenzione si vuole procedere con la pubblicazione a tariffa, " + vbCrLf + "l'operazione è inreversibile!!", vbQuestion + vbYesNo, "Conferma") = vbNo Then Exit Sub
    
    Anno = Left$(idSimulazione, Len(idSimulazione) - 2)
    
    'ParametriSistema
    gDBC.Execute ("Update " & _
                   "(select * from ParametriSistema as p  WHERE p.AnnoRiferimento='" & Anno & "' and p.codIstat='" & CodComune & "')  as T1 " & _
                   "inner join (select * from ParametriSistema as p1  WHERE p1.AnnoRiferimento='" & idSimulazione & "' and p1.codIstat='" & CodComune & "')  as T2 " & _
                   "on T1.codistat=T2.codistat and T1.parametro =T2.parametro " & _
                   "Set T1.Valore = T2.Valore")
    
    
    'CategorieAttivitaComunali
    gDBC.Execute ("UPDATE " & _
        "(SELECT * FROM CategorieAttivitaComunali WHERE CategorieAttivitaComunali.CodComune='" & CodComune & "' AND CategorieAttivitaComunali.AnnoRiferimento='" & Anno & "') AS T " & _
        "INNER JOIN (SELECT * FROM CategorieAttivitaComunali WHERE CategorieAttivitaComunali.CodComune='" & CodComune & "' AND CategorieAttivitaComunali.AnnoRiferimento='" & idSimulazione & "') AS T1 " & _
        "ON (T.IdCategoriaAttivita = T1.IdCategoriaAttivita) AND (T.CodComune = T1.CodComune) " & _
        "SET T.Superficie = [T1].[Superficie], T.NumRuoli = [T1].[NumRuoli], T.Tariffa = [T1].[Tariffa], T.TariffaPF = [T1].[TariffaPF], T.TariffaPV = [T1].[TariffaPV], T.SuperficieM = [T1].[SuperficieM]")

        
    'UtenzeDomesticheComunali
    gDBC.Execute ("update " & _
        "(SELECT * FROM UtenzeDomesticheComunali WHERE UtenzeDomesticheComunali.CodComune='" & CodComune & "'  AND UtenzeDomesticheComunali.AnnoRiferimento='" & Anno & "' ) as T " & _
        "Inner Join (SELECT * FROM UtenzeDomesticheComunali WHERE UtenzeDomesticheComunali.CodComune='" & CodComune & "'  AND UtenzeDomesticheComunali.AnnoRiferimento='" & idSimulazione & "' ) as T1 " & _
        "on T.CodComune= T1.CodComune and T.IdNumeriComponenti=T1.IdNumeriComponenti " & _
        "set   T.TariffaNComp=T1.TariffaNComp, T.TariffaSuperficie= T1.TariffaSuperficie,T.Superficie=T1.Superficie, T.SuperficieM=T1.SuperficieM")
    
    'ValoriCoefficienti507UDComunali
    gDBC.Execute ("update " & _
                  "(SELECT * FROM ValoriCoefficienti507UDComunali WHERE ValoriCoefficienti507UDComunali.CodComune='" & CodComune & "' AND ValoriCoefficienti507UDComunali.AnnoRiferimento='" & Anno & "') as T " & _
                  "Inner Join " & _
                  "(SELECT * FROM ValoriCoefficienti507UDComunali WHERE ValoriCoefficienti507UDComunali.CodComune='" & CodComune & "' AND ValoriCoefficienti507UDComunali.AnnoRiferimento='" & idSimulazione & "') as T1 " & _
                  "on T.CodComune = T1.CodComune and T.CodUtenza=T1.CodUtenza " & _
                  "set T.valore=T1.valore")
    'ValoriCoefficientiCAComunali
    gDBC.Execute ("update " & _
                  "(SELECT * FROM ValoriCoefficientiCAComunali WHERE ValoriCoefficientiCAComunali.CodComune='" & CodComune & "' AND ValoriCoefficientiCAComunali.AnnoRiferimento='" & Anno & "') AS T " & _
                  "Inner Join (SELECT * FROM ValoriCoefficientiCAComunali WHERE ValoriCoefficientiCAComunali.CodComune='" & CodComune & "' AND ValoriCoefficientiCAComunali.AnnoRiferimento='" & idSimulazione & "') AS T1 " & _
                  "on T.CodComune=T1.CodComune and  T.IdCoefficiente=T1.IdCoefficiente set T.Valore=T1.Valore")
    'ValoriCoefficientiIQSCAComunali
    gDBC.Execute ("update " & _
                "(SELECT * FROM ValoriCoefficientiIQSCAComunali WHERE ValoriCoefficientiIQSCAComunali.CodComune='" & CodComune & "' AND ValoriCoefficientiIQSCAComunali.AnnoRiferimento='" & Anno & "') as T " & _
                "Inner Join (SELECT * FROM ValoriCoefficientiIQSCAComunali WHERE ValoriCoefficientiIQSCAComunali.CodComune='" & CodComune & "' AND ValoriCoefficientiIQSCAComunali.AnnoRiferimento='" & idSimulazione & "') as T1 " & _
                "on T.CodComune=T1.CodComune and T.CodAttivita=T1.CodAttivita set T.IQS=T1.IQS,T.IPS=T1.IPS")
    'ValoriCoefficientiUDComunali
    gDBC.Execute ("update " & _
                  "(SELECT * FROM ValoriCoefficientiUDComunali WHERE ValoriCoefficientiUDComunali.CodComune='" & CodComune & "' AND ValoriCoefficientiUDComunali.AnnoRiferimento='" & Anno & "') as T " & _
                  "Inner Join (SELECT * FROM ValoriCoefficientiUDComunali WHERE ValoriCoefficientiUDComunali.CodComune='" & CodComune & "' AND ValoriCoefficientiUDComunali.AnnoRiferimento='" & idSimulazione & "') as T1 " & _
                  "on T.CodComune=T1.CodComune and  T.IdCoefficiente=T1.IdCoefficiente set T.valore= T1.valore")
    'DatiFinaliCalcoloTariffa
    gDBC.Execute ("update " & _
                "(SELECT * FROM DatiFinaliCalcoloTariffa WHERE DatiFinaliCalcoloTariffa.CodTipoTariffa=" & TipoMetodo & " and  DatiFinaliCalcoloTariffa.CodComune='" & CodComune & "' AND DatiFinaliCalcoloTariffa.AnnoRiferimento='" & Anno & "' ) as T " & _
                "Inner Join (SELECT * FROM DatiFinaliCalcoloTariffa WHERE DatiFinaliCalcoloTariffa.CodComune='" & CodComune & "' AND DatiFinaliCalcoloTariffa.AnnoRiferimento='" & idSimulazione & "' ) as T1 " & _
                "on T.CodComune=T1.CodComune " & _
                "set T.QVDOM=T1.QVDOM,T.QFDOM=T1.QFDOM,T.QUVDOM=T1.QUVDOM,T.QUFDOM=T1.QUFDOM, " & _
                "T.CUDOM=T1.CUDOM,T.QVNDOM=T1.QVNDOM,T.QFNDOM=T1.QFNDOM,T.CUNDOM=T1.CUNDOM,T.QAPFNDOM=T1.QAPFNDOM")
    'DatiFinaliCalcoloTarsu507
    gDBC.Execute ("update " & _
            "(SELECT * FROM DatiFinaliCalcoloTarsu507 WHERE DatiFinaliCalcoloTarsu507.CodComune='" & CodComune & "'  AND DatiFinaliCalcoloTarsu507.AnnoRiferimento='" & Anno & "' ) AS T " & _
            "Inner Join (SELECT * FROM DatiFinaliCalcoloTarsu507 WHERE DatiFinaliCalcoloTarsu507.CodComune='" & CodComune & "'  AND DatiFinaliCalcoloTarsu507.AnnoRiferimento='" & idSimulazione & "' ) AS T1 " & _
            "on  T.CodComune=T1.CodComune set  T.CostoTotale=T1.CostoTotale,T.Gettito=T1.Gettito," & _
            "T.CMG=T1.CMG,T.IQSUtenzaDomestica=T1.IQSUtenzaDomestica,T.IPSUtenzaDomestica=T1.IPSUtenzaDomestica,T.CoeffCorrettivoIPS=T1.CoeffCorrettivoIPS,T.ProdMedia=T1.ProdMedia,T.Copertura=T1.Copertura")
    'Conto Economico
    gDBC.Execute ("update " & _
            "(SELECT * FROM CostoComunale WHERE CostoComunale.CodComune='" & CodComune & "'  AND CostoComunale.AnnoRiferimento='" & Anno & "' ) AS T " & _
            "Inner Join (SELECT * FROM CostoComunale WHERE CostoComunale.CodComune='" & CodComune & "'  AND CostoComunale.AnnoRiferimento='" & idSimulazione & "' ) AS T1 " & _
            "on  T.CodComune=T1.CodComune " & _
                "SET T.Tot_Servizio = [T1].[Tot_Servizio],T.Gettito = [T1].[Gettito],T.CG = [T1].[CG], " & _
                "T.CSL = [T1].[CSL], T.CRT = [T1].[CRT],T.CTS = [T1].[CTS],T.AC = [T1].[AC], " & _
                "T.CGIND = [T1].[CGIND], T.CRD = [T1].[CRD],T.CTR_COSTI = [T1].[CTR_COSTI], " & _
                "T.CTR_RICAVI = [T1].[CTR_RICAVI],T.CTR = [T1].[CTR], T.CON = [T1].[CON], " & _
                "T.Y = [T1].[Y],  T.CGD = [T1].[CGD],T.CC = [T1].[CC],T.CARC = [T1].[CARC], " & _
                "T.CGG = [T1].[CGG], T.CCD = [T1].[CCD], T.CK = [T1].[CK],  T.AMM = [T1].[AMM], " & _
                "T.ACC = [T1].[ACC], T.R = [T1].[R],T.Rn = [T1].[RN], T.Kn_Anno_Prec = [T1].[KN_ANNO_PREC], " & _
                "T.Investimenti_Programmati = [T1].[Investimenti_Programmati], T.Fn = [T1].[FN],T.Tot_CSL = [T1].[Tot_CSL], " & _
                "T.Tot_CRT = [T1].[Tot_CRT],T.Tot_CRD = [T1].[Tot_CRD],T.Tot_CTS = [T1].[Tot_CTS], T.Tot_CTR = [T1].[Tot_CTR]")

    'DatiTecniciComune
    QuantitaRifiutiDaUtenzeAttivita CodComune, Anno
    
    gDBC.CommitTrans
    
    Exit Sub
PublicaComeTariffaERR:
    
    If bCommit Then gDBC.RollbackTrans
    
    MsgBox Error, vbCritical, App.Title
End Sub



