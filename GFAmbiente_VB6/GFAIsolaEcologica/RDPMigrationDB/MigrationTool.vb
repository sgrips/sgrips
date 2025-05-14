Imports System.Data.OleDb
Imports System.Text

Public Class Database

    Dim _dbPartenza As OleDb.OleDbConnection
    Dim _dbDestinazione As OleDb.OleDbConnection
    Dim _dbPartenzaName As String
    Dim _dbDestinazioneName As String
    Dim _TABLES_SKIP = "AssCodCerClassidiPericolo,CategoriaCER,SottocCategoriaCER,CodiciCER,Tariffe," + _
    "Comuni,Province,Regioni,ChartData,ChartDefault,Causali,ClassiPericolo,OperazioniSmaltimentoRecupero,Pese," + _
    "PeseEcoPuntoConCodCer,StatoFisicoCER,StatoLotto,StatoMovimento,TipiAnagrafe,UnitaDiMisura,tmpBadge,tmpImmobiliUtenti,tmpStampaRiferimentiRegistro," + _
    "tmpUtentiDaAggiornare"

    Property DestinazioneName() As String
        Get
            Return _dbDestinazioneName
        End Get
        Set(ByVal value As String)
            _dbDestinazioneName = value
        End Set
    End Property

    Property PartenzaName() As String
        Get
            Return _dbPartenzaName
        End Get
        Set(ByVal value As String)
            _dbPartenzaName = value
        End Set
    End Property

    Public Sub New()
        _TABLES_SKIP = Strings.UCase(_TABLES_SKIP)
    End Sub

    ReadOnly Property DBPartenza() As OleDb.OleDbConnection
        Get
            Return _dbPartenza
        End Get

    End Property

    ReadOnly Property DBDestinazione() As OleDb.OleDbConnection
        Get
            Return _dbDestinazione
        End Get

    End Property


    Public Sub ConnectionToStart()

        If _dbPartenza Is Nothing OrElse _dbPartenza.State = ConnectionState.Closed Then
            _dbPartenza = New OleDb.OleDbConnection
            Dim cntString = New System.Data.Common.DbConnectionStringBuilder
            cntString.ConnectionString = My.Settings.MDBConnection

            cntString("Data Source") = _dbPartenzaName
            cntString("JET Oledb:Database Password") = "gfa"

            _dbPartenza.ConnectionString = cntString.ConnectionString
            _dbPartenza.Open()

        End If


    End Sub

    Public Sub ConnectionToDestination()

        If _dbDestinazione Is Nothing OrElse _dbDestinazione.State = ConnectionState.Closed Then
            _dbDestinazione = New OleDb.OleDbConnection
            Dim cntString = New System.Data.Common.DbConnectionStringBuilder
            cntString.ConnectionString = My.Settings.MDBConnection

            cntString("Data Source") = _dbDestinazioneName
            cntString("JET Oledb:Database Password") = "gfa"

            _dbDestinazione.ConnectionString = cntString.ConnectionString
            _dbDestinazione.Open()

        End If

    End Sub


    Public Sub GenerateScript()
        Dim dtSetTableAdat As New ReportDataSetTableAdapters.TabelleNuoveTableAdapter
        Dim dataSetReport As New ReportDataSet

        Dim strBuff As New StringBuilder

        dtSetTableAdat.Connection.ConnectionString = My.Settings.ReportConnectionString
        dtSetTableAdat.Connection.Open()
        dtSetTableAdat.Fill(dataSetReport.TabelleNuove)

        For Each r As ReportDataSet.TabelleNuoveRow In dataSetReport.TabelleNuove.Rows
            strBuff.Append("CREATE TABLE" + r.Tabella)
            Dim schemaTableColumns As DataTable = _
                               _dbPartenza.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, _
                               New Object() {Nothing, Nothing, r.Tabella, Nothing})

            schemaTableColumns.WriteXml(Console.Out)

            Dim schemaTablePrimary_Keys As DataTable = _
                               _dbPartenza.GetOleDbSchemaTable(OleDbSchemaGuid.Primary_Keys, _
                               New Object() {Nothing, Nothing, r.Tabella})
            Console.WriteLine("------------------------")
            schemaTablePrimary_Keys.WriteXml(Console.Out)

            Dim schemaTable_Keys As DataTable = _
                               _dbPartenza.GetOleDbSchemaTable(OleDbSchemaGuid.Key_Column_Usage, _
                               New Object() {Nothing, Nothing, Nothing, Nothing, Nothing, r.Tabella, Nothing})
            Console.WriteLine("------------------------")
            schemaTable_Keys.WriteXml(Console.Out)
        Next
        dtSetTableAdat.Connection.Close()
    End Sub

    Public Sub AnalisyTables()
        'tabella nuove tabelle
        Dim _newTablesDataSet = New ReportDataSet
        Dim _tblAdpTables = New ReportDataSetTableAdapters.TabelleNuoveTableAdapter
        Dim _newTables = _newTablesDataSet.TabelleNuove

        Dim _tblAdpTablesColumns = New ReportDataSetTableAdapters.CampiNuoviTableAdapter
        Dim _newTablesColumns = _newTablesDataSet.CampiNuovi


        Dim conStrReport As Data.Common.DbConnectionStringBuilder = New Data.Common.DbConnectionStringBuilder
        conStrReport.ConnectionString = My.Settings.ReportConnectionString

        _tblAdpTables.Connection = New OleDb.OleDbConnection(conStrReport.ConnectionString)
        _tblAdpTables.Connection.Open()
        _tblAdpTables.ClearBeforeFill = True
        _tblAdpTables.DeleteAll()

        _tblAdpTables.Fill(_newTables)

        _tblAdpTablesColumns.Connection = New OleDb.OleDbConnection(conStrReport.ConnectionString)
        _tblAdpTablesColumns.Connection.Open()
        _tblAdpTablesColumns.ClearBeforeFill = True
        _tblAdpTablesColumns.DeleteAll()

        _tblAdpTables.Fill(_newTables)

        Dim schemaTableOrig As DataTable = _
            _dbPartenza.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, _
            New Object() {Nothing, Nothing, Nothing, "TABLE"})

        For Each row As DataRow In schemaTableOrig.Rows

            Dim tableName = Strings.UCase(row.Item("TABLE_NAME"))
            'per ogni tabella presente nel database di origine verifico la struttura con quello di destinazione
            If IsTableInDestinationDB(tableName, _dbDestinazione) Then
                If InStr(_TABLES_SKIP, tableName) = 0 Then
                    Dim schemaTableColumns As DataTable = _
                                _dbPartenza.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, _
                                New Object() {Nothing, Nothing, tableName, Nothing})

                    For i = 0 To schemaTableColumns.Rows.Count - 1
                        If Not IsColumnInDestinationDB(tableName, schemaTableColumns.Rows(i).Item("COLUMN_NAME"), _dbDestinazione) Then
                            'se non è presente la colonna la inserisco tra quelle nuove
                            Dim row1 = _newTablesColumns.NewCampiNuoviRow
                            row1.Tabella = tableName
                            row1.Campo = schemaTableColumns.Rows(i).Item("COLUMN_NAME")
                            _newTablesColumns.AddCampiNuoviRow(row1)
                        End If
                    Next

                End If
            Else
                'se non è presente la tabella la inserisco tra quelle nuove
                Dim row1 = _newTables.NewTabelleNuoveRow
                row1.Tabella = tableName
                _newTables.AddTabelleNuoveRow(row1)
            End If

        Next


        _tblAdpTables.Update(_newTables)
        _tblAdpTablesColumns.Update(_newTablesColumns)
        _newTablesDataSet.AcceptChanges()

        _newTablesDataSet.Dispose()
        _tblAdpTables.Connection.Close()

    End Sub

    Private Function IsTableInDestinationDB(ByVal tablename As String, ByVal destinationDB As OleDb.OleDbConnection)
        Dim schemaTable As DataTable = _
            destinationDB.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, _
            New Object() {Nothing, Nothing, tablename, "TABLE"})

        Return schemaTable.Rows.Count > 0
    End Function

    Private Function IsColumnInDestinationDB(ByVal tablename As String, ByVal columnname As String, ByVal destinationDB As OleDb.OleDbConnection)
        Dim schemaTableColumns As DataTable = _
                    destinationDB.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, _
                    New Object() {Nothing, Nothing, tablename, columnname})


        Return schemaTableColumns.Rows.Count > 0
    End Function

    Public Sub StartMigration()
        Dim schemaTable As DataTable = _
            _dbDestinazione.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, _
            New Object() {Nothing, Nothing, Nothing, "TABLE"})

        For Each row As DataRow In schemaTable.Rows
            Dim tableName = Strings.UCase(row.Item("TABLE_NAME"))

            'per ogni tabella presnte anche nel database di origine creo l'insert sql per inserire i dati in quello di destinazione
            Dim schemaTableOrig As DataTable = _
                _dbPartenza.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, _
                New Object() {Nothing, Nothing, tableName, "TABLE"})

            If schemaTableOrig.Rows.Count > 0 And InStr(_TABLES_SKIP, tableName) = 0 Then
                Dim cmd = New OleDb.OleDbCommand
                cmd.Connection = _dbDestinazione

                Dim schemaTableColumns As DataTable = _
                            _dbDestinazione.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, _
                            New Object() {Nothing, Nothing, tableName, Nothing})

                Dim arrValues() As OleDb.OleDbParameter = Nothing
                ReDim arrValues(schemaTableColumns.Rows.Count - 1)

                cmd.CommandText = "insert into {0} ({1}) values ({2})"
                Dim fieldNames = ""
                Dim fieldNamesQuestionMark = ""
                For i = 0 To schemaTableColumns.Rows.Count - 1
                    Dim col = schemaTableColumns.Rows(i)
                    Dim typeField As OleDbType = [Enum].Parse(GetType(OleDbType), col.Item("DATA_TYPE"))
                    Dim p = New OleDb.OleDbParameter

                    p.ParameterName = col.Item("COLUMN_NAME")
                    p.Value = DBNull.Value
                    p.OleDbType = typeField

                    fieldNames = fieldNames + p.ParameterName + If(i = schemaTableColumns.Rows.Count - 1, String.Empty, ",")
                    fieldNamesQuestionMark = fieldNamesQuestionMark + "?" + If(i = schemaTableColumns.Rows.Count - 1, String.Empty, ",")
                    arrValues(i) = p
                Next
                cmd.CommandText = String.Format(cmd.CommandText, tableName, fieldNames, fieldNamesQuestionMark)
                cmd.Parameters.AddRange(arrValues)

                'leggo i dati dalla tabella orgine e li scrivo su quella di destionazione
                Dim cmdSelectPartenza = _dbPartenza.CreateCommand()
                cmdSelectPartenza.CommandText = String.Format("select * from {0}", tableName)
                Dim oldDbDataRd As OleDb.OleDbDataReader = cmdSelectPartenza.ExecuteReader()

                If oldDbDataRd.HasRows Then
                    'se vi sono record li copio, altrimenti passo ad altra tabella
                    Form1.UpdateLblInfo(tableName)
                    Do While oldDbDataRd.Read
                        For i = 0 To oldDbDataRd.FieldCount - 1
                            If cmd.Parameters.Contains(oldDbDataRd.GetName(i)) Then
                                cmd.Parameters(oldDbDataRd.GetName(i)).Value = oldDbDataRd.Item(i)
                            End If
                        Next
                    Loop
                    cmd.ExecuteNonQuery()
                End If

                'schemaTableColumns.WriteXml("c:\temp\tableColumns.xml")

                oldDbDataRd.Close()
                schemaTableColumns.Dispose()

                cmd.Connection = Nothing
                cmd.Dispose()
            End If
            schemaTableOrig.Dispose()
        Next

    End Sub



End Class
