Imports System.IO
Imports System.Data.Common
Imports Microsoft.Win32
Imports ITAmbiente.easyrdc.Sync.DataSetMDB

Public Class SyncToDB

    Dim _dataSetBadge As DataSetMDBTableAdapters.BadgeTableAdapter
    Dim _dataSetUtenti As DataSetMDBTableAdapters.UtentiTableAdapter
    Dim _dataSetImmobili As DataSetMDBTableAdapters.ImmobiliTableAdapter
    Dim _dataSetMovimenti As DataSetMDBTableAdapters.MovimentiTableAdapter
    Dim _dataSetSocieta As DataSetMDBTableAdapters.SocietaTableAdapter
    Dim _dataSetSyncDB As DataSetMDBTableAdapters.SyncDBTableAdapter
    Dim _dataSetTotaliPerBadge As DataSetMDBTableAdapters.TotaliPerBadgeTableAdapter

    Dim _codiciCer As DataSetMDBTableAdapters.CodiciCerTableAdapter
    Dim _entityID As EntityID = New EntityID()
    Dim _clients As New DataTable("Clients")

    Dim _cmdTotalForBadge As New OleDb.OleDbCommand With {.CommandText = "INSERT INTO TotaliPerBadge ( CodBadge, CodUtente, CodiciCerCodCategoria, CodiciCerCodSottoCategoria, CodiciCerCodCer, UnitaDiMisura, SomBonus, PuntiBonus, SomQty ) SELECT Badge.CodBadge, Movimenti.CodUtente, Movimenti.CodCategoria, Movimenti.CodSottoCategoria, " + _
                    "Movimenti.CodCer,UnitaMisura,Sum(IIf([segno]=1,[BonusDovuto],-[BonusDovuto])) AS SomBonus,Sum(IIf([segno]=1,[PuntiDovuti],-[PuntiDovuti])) AS PuntiBonus, Sum(IIf([segno]=1,[quantita],-[quantita])) AS SomQty " + _
                    "FROM (Movimenti INNER JOIN Badge ON Movimenti.CodUtente = Badge.CodUtente) INNER JOIN Causali ON Movimenti.CodCausale = Causali.CodCausale " + _
                    "WHERE Movimenti.DataMovimento Between ? And ?  AND Movimenti.stato<>'2'" + _
                    "GROUP BY Badge.CodBadge, Movimenti.CodUtente, Movimenti.CodCategoria, Movimenti.CodSottoCategoria, Movimenti.CodCer,UnitaMisura", .CommandType = CommandType.Text}

    Dim _cmdCountFilesToSend As New OleDb.OleDbCommand With {.CommandText = "Select ToSend from IsAllFilesWereSync "}


    Dim _codComune As String
    Dim _codCausale As String
    Dim _flgGeneraGiacenza As Boolean
    Dim _mainForm As Setting

    WriteOnly Property MainForm()
        Set(ByVal value)
            _mainForm = value
        End Set
    End Property


    Public Sub New()
        _dataSetBadge = New DataSetMDBTableAdapters.BadgeTableAdapter
        _dataSetUtenti = New DataSetMDBTableAdapters.UtentiTableAdapter
        _dataSetImmobili = New DataSetMDBTableAdapters.ImmobiliTableAdapter
        _dataSetMovimenti = New DataSetMDBTableAdapters.MovimentiTableAdapter
        _dataSetSocieta = New DataSetMDBTableAdapters.SocietaTableAdapter
        _dataSetSyncDB = New DataSetMDBTableAdapters.SyncDBTableAdapter
        _dataSetTotaliPerBadge = New DataSetMDBTableAdapters.TotaliPerBadgeTableAdapter
        _codiciCer = New DataSetMDBTableAdapters.CodiciCerTableAdapter

        RefreshConnection()

        AccociateConnectionToDataSet()

        Dim p As OleDb.OleDbParameter = New OleDb.OleDbParameter("DataIni", OleDb.OleDbType.Date)
        _cmdTotalForBadge.Parameters.Add(p)
        p = New OleDb.OleDbParameter("DataFine", OleDb.OleDbType.Date)
        _cmdTotalForBadge.Parameters.Add(p)


        p = New OleDb.OleDbParameter("NameFile0", OleDb.OleDbType.VarChar, 255)
        _cmdCountFilesToSend.Parameters.Add(p)
        p = New OleDb.OleDbParameter("NameFile1", OleDb.OleDbType.VarChar, 255)
        _cmdCountFilesToSend.Parameters.Add(p)


        _clients.Columns.Add("CodPesa", GetType(String))
        _clients.Columns.Add("PcName", GetType(String))



    End Sub

    Public Sub GetListClients()
        Dim rs As DbDataReader
        Dim cmd As DbCommand
        cmd = Util.Connection.ConnectionMDB.CreateCommand

        cmd.CommandText = "Select CodPesa,PcName from Pese where PcName<>'' and not PcName is null"

        If cmd.Connection Is Nothing OrElse cmd.Connection.State = ConnectionState.Closed Then Exit Sub

        rs = cmd.ExecuteReader

        _clients.Clear()
        _clients.Load(rs)
    End Sub

    ReadOnly Property Clients() As DataTable
        Get
            GetListClients()

            Return _clients
        End Get
    End Property


    Public Sub RefreshConnection()

        Try
            _dataSetBadge.Dispose()
            _dataSetUtenti.Dispose()
            _dataSetImmobili.Dispose()
            _codiciCer.Dispose()
            _dataSetMovimenti.Dispose()
            _dataSetSocieta.Dispose()
            _dataSetSyncDB.Dispose()
            _dataSetTotaliPerBadge.Dispose()

            Util.Connection.CloseMDB()

            _codComune = String.Empty

        Catch ex As Exception
            My.Application.Log.WriteException(ex)
        End Try
    End Sub

    Private Sub AccociateConnectionToDataSet()
        _dataSetBadge.Connection = Util.Connection.ConnectionMDB
        _dataSetUtenti.Connection = Util.Connection.ConnectionMDB
        _dataSetImmobili.Connection = Util.Connection.ConnectionMDB
        _codiciCer.Connection = Util.Connection.ConnectionMDB
        _dataSetMovimenti.Connection = Util.Connection.ConnectionMDB
        _dataSetSocieta.Connection = Util.Connection.ConnectionMDB
        _dataSetSyncDB.Connection = Util.Connection.ConnectionMDB
        _dataSetTotaliPerBadge.Connection = Util.Connection.ConnectionMDB
        
    End Sub


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ImportOldFileCSV()
        Dim q = From f In System.IO.Directory.GetFiles(My.Settings.FolderNameSync, "*.csv")

        For Each f In q
            StartImport(f)
        Next
    End Sub


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub RenameFileEvent(ByVal sender As Object, ByVal e As IO.RenamedEventArgs)
        Dim sendFSW = DirectCast(sender, FileSystemWatcher)
        StartImport(e.FullPath)
    End Sub


    Private Sub BounsAttivoEPremioAttivo(ByVal codUtente As Long, ByRef bonus As Boolean, ByRef premio As Boolean, ByVal dbTransaction As DbTransaction)
        Dim rs As OleDb.OleDbDataReader
        Dim cmd As OleDb.OleDbCommand


        cmd = Util.Connection.ConnectionMDB.CreateCommand
        cmd.Transaction = dbTransaction
        cmd.CommandText = "SELECT TipiAnagrafe.AttivoBonus, TipiAnagrafe.AttivoPremio " & _
                              "FROM TipiAnagrafe INNER JOIN Utenti ON TipiAnagrafe.TipoAnagrafe = Utenti.TipoAnagrafe " & _
                              "WHERE Utenti.CodUtente=" & codUtente
        rs = cmd.ExecuteReader()

        If rs.Read Then
            bonus = rs.GetBoolean(0)
            premio = rs.GetBoolean(1)
        End If
    End Sub

    Dim objLockImport = New Object
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="fileName"></param>
    ''' <remarks></remarks>
    Private Sub StartImport(ByVal fileName As String)
        SyncLock objLockImport
            Dim cmdGetMov As New Data.OleDb.OleDbCommand
            Dim dataReaderMov As DbDataReader = Nothing
            Dim fileInfo As IO.FileInfo = My.Computer.FileSystem.GetFileInfo(fileName)
            Dim dbTransaction As OleDb.OleDbTransaction = Nothing
            Dim codPesa As String = vbNullString
            Dim beginTrans As Boolean = False
            Dim bErr As Boolean = False

            Dim sRifMov As String
            Dim cmd = Util.Connection.ConnectionMDB.CreateCommand()
            cmd.CommandText = "select * from movimenti"



            'se la dimensione del file è maggiore di 0 allora inizio il trasferimento
            If fileInfo.Length > 0 And fileInfo.Extension.ToUpper = ".CSV" Then

                Threading.Thread.Sleep(200)

                Try
                    AccociateConnectionToDataSet()

                    If String.IsNullOrEmpty(_codComune) Then

                        Dim codSocieta = ITAmbiente.easyrdc.Bridge.SettingsRDC.Instance.GetPropertyValue(ITAmbiente.easyrdc.Bridge.SettingsRDC.SETTING_SOCIETA)
                        Dim socDataTable As SocietaDataTable = _dataSetSocieta.GetDataBy(codSocieta)

                        _dataSetSocieta.Connection = Util.Connection.ConnectionMDB
                        If socDataTable.Count > 0 Then
                            _codComune = socDataTable.Item(0).CodComune
                        End If

                    End If

                    _codCausale = ITAmbiente.easyrdc.Bridge.SettingsRDC.Instance.GetPropertyValue(ITAmbiente.easyrdc.Bridge.SettingsRDC.SETTING_ECOPUNTO_CAUSALE_MOVIMENTO_IN)
                    _flgGeneraGiacenza = ITAmbiente.easyrdc.Bridge.SettingsRDC.Instance.GetPropertyValue(ITAmbiente.easyrdc.Bridge.SettingsRDC.SETTING_ECOPUNTO_GENERA_GIACENZA)


                    cmdGetMov.CommandText = "select * from [" + fileInfo.Name + "]"
                    cmdGetMov.Connection = Util.Connection.ConnectionTxt("Movimenti.csv", fileInfo.FullName)
                    dataReaderMov = cmdGetMov.ExecuteReader

                    dbTransaction = Util.Connection.ConnectionMDB.BeginTransaction
                    beginTrans = True
                    Do While dataReaderMov.Read()
                        'creazione utente, creazione immobile, creazione badge
                        Dim codUtente
                        Dim codImmobile
                        Dim codComune As String = If(dataReaderMov.Item("CodIstat") = 0, _codComune, dataReaderMov.Item("CodIstat"))

                        If dataReaderMov.Item("codUtente") Is DBNull.Value OrElse dataReaderMov.Item("codUtente") = 0 Then
                            Dim arrList = CreateBadge(dataReaderMov.Item("codbadge"), codComune, dbTransaction)
                            codUtente = arrList.Item(0)
                            codImmobile = arrList.Item(1)
                        Else
                            codUtente = dataReaderMov.Item("codUtente")
                            codImmobile = dataReaderMov.Item("codImmobile")
                        End If


                        'creazione movimento
                        sRifMov = dataReaderMov.Item("ID_CodPesa")
                        codPesa = Right(sRifMov, 2)
                        Dim BonusDovuto As Double
                        Dim BonusPunti As Double
                        Dim codCerTable As New CodiciCerDataTable

                        _codiciCer.Connection = Util.Connection.ConnectionMDB
                        _codiciCer.Transaction = dbTransaction
                        _codiciCer.FillByCer(codCerTable, codComune, dataReaderMov.Item("CodCategoria"), dataReaderMov.Item("CodSottoCategoria"), _
                                             dataReaderMov.Item("CodCer"), dataReaderMov.Item("CodSubCer"))

                        Dim bonusAttivo As Boolean = False
                        Dim premioAttivo As Boolean = False

                        BounsAttivoEPremioAttivo(codUtente, bonusAttivo, premioAttivo, dbTransaction)

                        BonusDovuto = 0
                        BonusPunti = 0

                        If bonusAttivo Then BonusDovuto = codCerTable.Item(0).CorrispettivoBonus * dataReaderMov.Item("Quantita")
                        If premioAttivo Then BonusPunti = codCerTable.Item(0).BonusPunti * dataReaderMov.Item("Quantita")

                        CreateMovement(_entityID.GetNewEntityID("MOVIMENTI", dbTransaction), codImmobile, codUtente, dataReaderMov.Item("CodCategoria"), dataReaderMov.Item("codSottoCategoria"), _
                                       dataReaderMov.Item("CodCer"), dataReaderMov.Item("CodSubCer"), codComune, dataReaderMov.Item("Data"), dataReaderMov.Item("Quantita"), 0, False, "Admin", BonusDovuto, BonusPunti, _
                                       dataReaderMov.Item("UnitaDiMisura"), Nothing, Nothing, Nothing, False, Nothing, Nothing, _
                                       _codCausale, Not _flgGeneraGiacenza, codPesa, sRifMov, dataReaderMov.Item("idAlibiMemory"), dbTransaction)

                    Loop

                    sRifMov = String.Empty
                    _entityID.SaveAll(dbTransaction)
                    dbTransaction.Commit()
                    beginTrans = False

                Catch ex As Exception
                    If beginTrans = True Then dbTransaction.Rollback()
                    My.Application.Log.WriteException(ex)
                    bErr = True

                Finally
                    cmdGetMov.Dispose()
                    Util.Connection.CloseTxt()
                    Util.Connection.RelaseConnectionMDB()

                    If dataReaderMov IsNot Nothing Then dataReaderMov.Close()

                    CloseDataSet()

                    'sposto in directory ERR
                    If bErr = True Then
                        If Not My.Computer.FileSystem.DirectoryExists(fileInfo.DirectoryName + "\ERR") Then
                            My.Computer.FileSystem.CreateDirectory(fileInfo.DirectoryName + "\ERR")
                        End If
                        My.Computer.FileSystem.MoveFile(fileInfo.FullName, fileInfo.DirectoryName + "\ERR\" + Path.GetFileNameWithoutExtension(fileInfo.Name) + "_" + Now.ToString("yyyyMMdd_hhmmss") + Path.GetExtension(fileInfo.Name), True)
                    Else
                        'sposto in directory WORK
                        If Not My.Computer.FileSystem.DirectoryExists(fileInfo.DirectoryName + "\WRK") Then
                            My.Computer.FileSystem.CreateDirectory(fileInfo.DirectoryName + "\WRK")
                        End If
                        My.Computer.FileSystem.MoveFile(fileInfo.FullName, fileInfo.DirectoryName + "\WRK\" + Path.GetFileNameWithoutExtension(fileInfo.Name) + "_" + Now.ToString("yyyyMMdd_hhmmss") + Path.GetExtension(fileInfo.Name), True)
                    End If

                End Try
            ElseIf fileInfo.Length > 0 And fileInfo.Extension.ToUpper = ".RESULT" Then
                'Movimenti+[ClientID]+[FlgError]+[FileName]
                Dim sFile As String = fileInfo.Name.Remove(0, Len("Movimenti"))
                Dim clientID As String = Left$(sFile, 2)
                bErr = Mid$(sFile, 2, 1) = 1

                fileName = Mid$(sFile, 4).Replace(".result", String.Empty)

                Threading.Thread.Sleep(300)
                My.Computer.FileSystem.DeleteFile(fileInfo.FullName)
                RemoveFileToSync(clientID, fileName)
            End If
        End SyncLock
    End Sub

    Private Sub CloseDataSet()
        _dataSetBadge.Connection.Close()
        _dataSetUtenti.Connection.Close()
        _dataSetImmobili.Connection.Close()
        _codiciCer.Connection.Close()
        _dataSetMovimenti.Connection.Close()
        _dataSetSocieta.Connection.Close()
        _dataSetSyncDB.Connection.Close()
    End Sub


    Public Sub CreateTotal()
        'scrive file con totali
        Dim _fileNameXMLTot As String = "MovimentiTotal"

        If My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CombinePath(My.Settings.FolderNameSync, _fileNameXMLTot) + ".XML") Then
            My.Computer.FileSystem.DeleteFile(My.Computer.FileSystem.CombinePath(My.Settings.FolderNameSync, _fileNameXMLTot) + ".XML")
        End If

        Dim firtDayOfYear = DateAdd(DateInterval.Second, +3, DateAdd(DateInterval.Year, DateDiff(DateInterval.Year, Date.MinValue, Today()), Date.MinValue))
        Dim lastDayOfYear = DateAdd(DateInterval.Second, -3, DateAdd(DateInterval.Year, DateDiff(DateInterval.Year, Date.MinValue, Today()) + 1, Date.MinValue))

        Dim cmd As OleDb.OleDbCommand = Util.Connection.ConnectionMDB.CreateCommand
        cmd.CommandText = "delete * from TotaliPerBadge"
        cmd.ExecuteNonQuery()


        _cmdTotalForBadge.Parameters(0).Value = firtDayOfYear
        _cmdTotalForBadge.Parameters(1).Value = lastDayOfYear
        _cmdTotalForBadge.Connection = Util.Connection.ConnectionMDB
        _cmdTotalForBadge.ExecuteNonQuery()

        'rilascio risorse 
        cmd.Dispose()

        Dim TotaliPerBadge As TotaliPerBadgeDataTable = New TotaliPerBadgeDataTable
        _dataSetTotaliPerBadge.Fill(TotaliPerBadge)

        TotaliPerBadge.WriteXml(My.Computer.FileSystem.CombinePath(My.Settings.FolderNameSync, _fileNameXMLTot) + ".TMP")
        My.Computer.FileSystem.RenameFile(My.Computer.FileSystem.CombinePath(My.Settings.FolderNameSync, _fileNameXMLTot) + ".TMP", _
                                           _fileNameXMLTot + ".XML")

        TotaliPerBadge.Dispose()
    End Sub

    ''' <summary>
    ''' Prepara i file per l'esportazione verso i clients
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub PrepareToSync()

        'prendo la lista dei clients e preparo i file da spedire
        GetListClients()

        For Each e In My.Computer.FileSystem.GetFiles(My.Settings.FolderNameSync, FileIO.SearchOption.SearchTopLevelOnly, "*.xml")
            Dim sfile = Path.GetFileName(e)
            For Each r As DataRow In _clients.Rows
                If _dataSetSyncDB.GetDataByClientIDNameFile(r.Item("CodPesa"), sfile).Count = 0 Then
                    _dataSetSyncDB.Insert(r.Item("CodPesa"), sfile, Now, False)
                Else
                    _dataSetSyncDB.Update(r.Item("CodPesa"), sfile, Now, False, r.Item("CodPesa"))
                End If
            Next

            'sposto il file da trasferire
            Dim sSendDir = My.Computer.FileSystem.CombinePath(My.Settings.FolderNameSync, "ToSend")
            If Not My.Computer.FileSystem.DirectoryExists(sSendDir) Then My.Computer.FileSystem.CreateDirectory(sSendDir)

            'se il file è già presente non posso renderlo disponibile per il trasferimento
            If Not My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CombinePath(sSendDir, sfile)) Then
                My.Computer.FileSystem.MoveFile(My.Computer.FileSystem.CombinePath(My.Settings.FolderNameSync, sfile), _
                                                My.Computer.FileSystem.CombinePath(sSendDir, sfile))
            End If
        Next

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Dim lockFileSend As New Object
    Public Sub SendFileToSync(ByVal clientID As String)
        SyncLock lockFileSend
            Dim syncTable As SyncDBDataTable = _dataSetSyncDB.GetDataByClientID(clientID)

            If syncTable.Count > 0 Then
                For Each r In syncTable
                    Dim pcName = r.Item("PCNAME")
                    Dim ipAdress = r.Item("IPADDRESS")
                    Dim remoteFile = String.Format("\\{0}\{1}\{2}", pcName, "SyncDB", Path.GetFileNameWithoutExtension(r.Item("NameFile")))



                    If Not r.Item("sync") Then
                        'trasferisco il file non sincronizzato dalla directory toSend al percorso remoto
                        Dim sFile = My.Computer.FileSystem.CombinePath(My.Settings.FolderNameSync + "\ToSend", r.Item("NameFile"))
                        Try
                            My.Computer.FileSystem.CopyFile(sFile, remoteFile + ".sync", True)
                            Threading.Thread.Sleep(300)
                            My.Computer.FileSystem.RenameFile(remoteFile + ".sync", Path.GetFileNameWithoutExtension(r.Item("NameFile")) + ".xml")

                            'invio il messaggio di file inviato solo nel caso di connessione PIPE
                            If My.Settings.ConnectionMode = ConnectionModeEnum.PIPE Then
                                _mainForm.SendMessage(clientID, _
                                    Framework.IO.Communication.Pipe.MessagePipeIDEnum.MessagePipeClientInfo, _
                                    r.Item("NameFile"), _
                                    Pipes.Messagecode.Constants.MESSAGE_CODE_FILESYNC)
                            End If
                        Catch ex As Exception
                            My.Application.Log.WriteException(ex)
                        End Try
                    End If
                Next
            End If
        End SyncLock
    End Sub

    Dim lockFileRemove As New Object
    Public Sub RemoveFileToSync(ByVal clientID As String, ByVal sFile As String)
        'imposto il file come spedito, poi se il file è stato inviato a tutti i clients lo rimuovo
        SyncLock lockFileRemove
            Dim syncTable As SyncDBDataTable = _dataSetSyncDB.GetDataByClientIDNameFile(clientID, sFile)

            If syncTable.Count > 0 Then
                Dim r = syncTable.Rows(0)
                If Not r.Item("sync") Then r.Item("sync") = True

                Dim sFileSent = My.Computer.FileSystem.CombinePath(My.Settings.FolderNameSync + "\ToSend", r.Item("NameFile"))
                Try
                    _dataSetSyncDB.Update(syncTable)
                    _cmdCountFilesToSend.Parameters(0).Value = r.Item("NameFile")
                    _cmdCountFilesToSend.Connection = Util.Connection.ConnectionMDB
                    Dim totToSend = _cmdCountFilesToSend.ExecuteScalar

                    If Not String.IsNullOrEmpty(sFileSent) And totToSend = 0 Then
                        My.Computer.FileSystem.DeleteFile(sFileSent)
                        _dataSetSyncDB.DeleteByNameFile(sFile)
                    End If

                Catch ex As Exception
                    My.Application.Log.WriteException(ex)
                End Try
            Else
                My.Application.Log.WriteEntry(String.Format("Attenzione non trovato clientID {0} con File {1}", clientID, sFile))
            End If
        End SyncLock
    End Sub



    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CreateBadge(ByVal codBadge As String, ByVal codComune As String, ByVal dbTransaction As OleDb.OleDbTransaction) As ArrayList
        Dim dataTable As New BadgeDataTable
        Dim arrList = New ArrayList(2)



        'se il badge non è presente lo inserisco
        _dataSetBadge.Connection = Util.Connection.ConnectionMDB
        _dataSetBadge.Transaction = dbTransaction

        _dataSetBadge.FillBy(dataTable, codBadge)

        If dataTable.Rows.Count = 0 Then
            'genero un utente fittizio
            Dim codUtente As Integer = _entityID.GetNewEntityID("UTENTI", dbTransaction)
            Dim codImmobile As Integer = -1

            If codUtente > 0 Then
                _dataSetUtenti.Connection = Util.Connection.ConnectionMDB
                _dataSetUtenti.Transaction = dbTransaction
                _dataSetUtenti.Insert(codComune, codUtente, codUtente, "F", vbNullString, ".", "Utente" + CStr(codUtente), "M", New DateTime, vbNullString, _
                                     ".", vbNullString, 0, vbNullString, codComune, 0, ".", 0, vbNullString, vbNullString, _
                                    vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, 0, vbNullString, 0, vbNullString, vbNullString, _
                                    vbNullString, vbNullString, vbNullString, vbNullString)

                'inserisco l'immobile
                codImmobile = _entityID.GetNewEntityID("IMMOBILI", dbTransaction)
                If codImmobile > 0 Then
                    _dataSetImmobili.Connection = Util.Connection.ConnectionMDB
                    _dataSetImmobili.Transaction = dbTransaction
                    _dataSetImmobili.Insert(codImmobile, codUtente, "P", 0, codComune, 0, ".", 0, vbNullString, vbNullString, vbNullString, _
                                            vbNullString, 0, 0, "Ud 00", "D", True, Now.Date, New DateTime(9999, 1, 1))

                    'inserisco il badge
                    _dataSetBadge.Connection = Util.Connection.ConnectionMDB
                    _dataSetBadge.Transaction = dbTransaction
                    _dataSetBadge.Insert(codBadge, codUtente, "F", 0, Now.Date, New DateTime(9999, 1, 1), codComune, codImmobile, "D", True)

                Else
                    Throw New Exception("Attenzione non è stato possibile prendere un numeratore IMMOBILI")
                End If
            Else
                Throw New Exception("Attenzione non è stato possibile prendere un numeratore UTENTI")
            End If


            arrList.Add(codUtente)
            arrList.Add(codImmobile)
        Else
            arrList.Add(dataTable.Rows(0).Item("CodUtente"))
            arrList.Add(dataTable.Rows(0).Item("CodImmobile"))
        End If

        Return arrList
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="nrMovimento"></param>
    ''' <param name="codImmobile"></param>
    ''' <param name="codUtente"></param>
    ''' <param name="codCategoria"></param>
    ''' <param name="codSottoCategoria"></param>
    ''' <param name="codCer"></param>
    ''' <param name="dataMovimento"></param>
    ''' <param name="quantita"></param>
    ''' <param name="Stato"></param>
    ''' <param name="manuale"></param>
    ''' <param name="codUser"></param>
    ''' <param name="bonusDovuto"></param>
    ''' <param name="unitaMisura"></param>
    ''' <param name="Targa"></param>
    ''' <param name="numeroFormulario"></param>
    ''' <param name="dataFormulario"></param>
    ''' <param name="flgForzatura"></param>
    ''' <param name="targaTrasportatore"></param>
    ''' <param name="codTrasportatore"></param>
    ''' <param name="codCausale"></param>
    ''' <param name="flgRegistrato"></param>
    ''' <param name="codpesa"></param>
    ''' <param name="rifMovImportato"></param>
    ''' <remarks></remarks>
    Private Sub CreateMovement(ByVal nrMovimento As Integer, ByVal codImmobile As Integer, ByVal codUtente As Integer, ByVal codCategoria As String, ByVal codSottoCategoria As String, ByVal codCer As String, _
            ByVal codSubCer As String, ByVal codComune As String, ByVal dataMovimento As DateTime, ByVal quantita As Double, ByVal Stato As String, ByVal manuale As Boolean, ByVal codUser As String, _
            ByVal bonusDovuto As Double, ByVal PuntiDovuti As Double, ByVal unitaMisura As String, ByVal Targa As String, ByVal numeroFormulario As String, ByVal dataFormulario As Global.System.Nullable(Of Date), _
            ByVal flgForzatura As Boolean, ByVal targaTrasportatore As String, ByVal codTrasportatore As Integer, _
            ByVal codCausale As String, ByVal flgRegistrato As Boolean, ByVal codpesa As String, ByVal rifMovImportato As String, ByVal idAlibiMemory As String, ByVal dbTransaction As OleDb.OleDbTransaction)

        Dim _movDataTable As New MovimentiDataTable


        _dataSetMovimenti.Connection = Util.Connection.ConnectionMDB
        _dataSetMovimenti.Transaction = dbTransaction
        _dataSetMovimenti.FillByRifMovImportato(_movDataTable, rifMovImportato)
        'attenzione la colonna quantità va impostata in double
        If _movDataTable.Count = 0 Then
            _dataSetMovimenti.Connection = Util.Connection.ConnectionMDB
            _dataSetMovimenti.Transaction = dbTransaction
            Try
                _dataSetMovimenti.Insert(codComune, nrMovimento, codImmobile, codUtente, Targa, codCategoria, codSottoCategoria, codCer, codSubCer, _
                                         codpesa, dataMovimento, quantita, unitaMisura, bonusDovuto, PuntiDovuti, targaTrasportatore, codTrasportatore, numeroFormulario, _
                                         dataFormulario, Stato, manuale, codCausale, flgForzatura, flgRegistrato, Now, codUser, rifMovImportato, 0, idAlibiMemory)

            Catch ex As Exception
                Throw New Exception(String.Format("Errore nel movimento idRif:{0}", rifMovImportato), ex)
            End Try


        Else
            Throw New Exception("Movimento idRif:" + rifMovImportato + " già presente")
        End If

    End Sub


    Protected Overrides Sub Finalize()

        _entityID.Close()

        _dataSetBadge.Dispose()
        _dataSetUtenti.Dispose()
        _dataSetImmobili.Dispose()
        _codiciCer.Dispose()
        _dataSetMovimenti.Dispose()
        _dataSetSocieta.Dispose()
        _dataSetSyncDB.Dispose()

        MyBase.Finalize()
    End Sub
End Class
