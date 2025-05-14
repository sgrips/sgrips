Imports System.Data.Common
Imports System.IO
Imports System.Security.Permissions
Imports System.Security
Imports System.Text
Imports ITAmbiente.easyrdc.Bridge


Namespace Util

    Friend NotInheritable Class Connection
        Private Shared _conTXT As DbConnection
        Private Shared _conMDB As OleDb.OleDbConnection
        Private Shared _keyFileSchemaLast As String


        Private Sub New()
            'not is instance
        End Sub

#Region "connessione file MDB"
        Public Shared ReadOnly Property ConnectionMDB() As OleDb.OleDbConnection
            Get
                If _conMDB Is Nothing Then
                    _conMDB = New OleDb.OleDbConnection

                    Dim conMDBString As New OleDb.OleDbConnectionStringBuilder
                    Try
                        conMDBString.ConnectionString = My.Settings.OleDBMSAccess
                        conMDBString.Item("Data Source") = My.Computer.FileSystem.CombinePath(My.Settings.PathDataBase, "IsolaEcologica.mdb")
                        conMDBString.Item("Jet OLEDB:Database Password") = Util.Security.Cryptography.Decrypt(My.Settings.PasswordDB, "ITAmbiente")

                        _conMDB.ConnectionString = conMDBString.ConnectionString
                        _conMDB.Open()
                    Catch ex As Exception
                        My.Application.Log.WriteException(ex)
                    End Try
                ElseIf _conMDB.State = ConnectionState.Closed Then
                    Try
                        _conMDB.Open()
                    Catch ex As Exception
                        My.Application.Log.WriteException(ex)
                    End Try
                End If

                Return _conMDB

            End Get
        End Property

        Public Shared Sub RelaseConnectionMDB()
            Try
                If _conMDB IsNot Nothing AndAlso _conMDB.State = ConnectionState.Open Then
                    _conMDB.Close()
                End If

            Catch ex As Exception
                My.Application.Log.WriteException(ex)
            End Try


        End Sub


        ''' <summary>
        ''' 
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub CloseMDB()
            'close File MDB
            Try
                My.Application.Log.WriteEntry("Chiusura della connessione File MDB in corso", TraceEventType.Verbose)
                If _conMDB IsNot Nothing AndAlso _conMDB.State = ConnectionState.Open Then
                    _conMDB.Close()
                    _conMDB.Dispose()
                    _conMDB = Nothing
                End If
                My.Application.Log.WriteEntry("Connessione File MDB chiusa", TraceEventType.Verbose)

            Catch ex As Exception
                My.Application.Log.WriteException(ex)
            End Try
        End Sub

#End Region


#Region "connessione file testo"

        Public Shared ReadOnly Property ConnectionTxt(ByVal keySchemaINI As String, ByVal fileTxt As String) As DbConnection
            Get
                Dim pathSchemaIni = Path.GetDirectoryName(fileTxt)
                Dim keyFileSchema As String = Path.GetFileName(fileTxt)

                If _conTXT Is Nothing OrElse _
                    _conTXT.ConnectionString.IndexOf(Path.GetDirectoryName(fileTxt)) = -1 OrElse _
                    _keyFileSchemaLast <> keyFileSchema Then

                    _keyFileSchemaLast = keyFileSchema
                    If CreateSchemaINI(keySchemaINI, keyFileSchema, pathSchemaIni) Then
                        Dim mConnessione As String = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                "Data Source=" & pathSchemaIni & ";" & _
                                "Extended Properties=""text;HDR=Yes;FMT=Delimiter"""
                        Try
                            'chiudo la precedente connessione se aperta
                            CloseTxt()
                            'creo una nuova connessione 
                            _conTXT = New OleDb.OleDbConnection
                            _conTXT.ConnectionString = mConnessione
                            _conTXT.Open()
                        Catch ex As Exception
                            'Imposto la connessione a nothing se si sono avuti errori 
                            My.Application.Log.WriteException(ex)
                            _conTXT = Nothing
                        End Try
                    End If
                End If
                Return _conTXT
            End Get
        End Property

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub CloseTxt()
            'close File Testo connection
            Try
                My.Application.Log.WriteEntry("Chiusura della connessione File Testo in corso", TraceEventType.Verbose)
                If _conTXT IsNot Nothing AndAlso _conTXT.State = ConnectionState.Open Then
                    _conTXT.Close()
                    _conTXT.Dispose()
                    _conTXT = Nothing
                End If
                My.Application.Log.WriteEntry("Connessione File Testo chiusa", TraceEventType.Verbose)

            Catch ex As Exception
                My.Application.Log.WriteException(ex)
            End Try
        End Sub

        ''' <summary>
        ''' Function per la generazione del file Schema.ini per l'accesso ai file testo con specifica SQL
        ''' </summary>
        ''' <param name="keySchemaINI"></param> 
        ''' <param name="pathSchemaIni"></param> 
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function CreateSchemaINI(ByVal keySchemaINI As String, ByVal keyFileSchema As String, ByVal pathSchemaIni As String) As Boolean
            Dim txt As System.IO.TextReader

            txt = New StringReader(My.Resources.SchemiINI.ToString)
            Dim xmlSchemaINI = XDocument.Load(txt)

            Dim xmlN = From e In xmlSchemaINI...<section> _
                       Select e _
                       Where (e.@name = keySchemaINI)



            If xmlN.Count > 0 Then
                'trovata la sezione genero lo schema.ini corrispondete seguendo le specifiche di microsoft
                '[FILENAME]
                'ColNameHeader = True 
                'Format =Delimited(";")
                'CharacterSet=ANSI
                'DecimalSymbol=,
                'TextDelimeter = " 
                'DateTimeFormat=yyyy/mm/dd
                Dim fileStream As FileStream = Nothing
                Dim TextWriter = Nothing
                Try
                    'verifico se l'utente ha l'autorizzazione in scrittura
                    Dim fIOPermission = New FileIOPermission(FileIOPermissionAccess.Write, CType(pathSchemaIni + "\Schema.ini", String))
                    fIOPermission.Demand()

                    fileStream = New FileStream(pathSchemaIni + "\Schema.ini", FileMode.Create)
                    TextWriter = New StreamWriter(fileStream)

                    Dim xmlElement = xmlN...<element>

                    'TextWriter.WriteLine("[" + Path.GetFileName(fName) + "]")
                    TextWriter.WriteLine("[" + keyFileSchema + "]")
                    TextWriter.WriteLine("ColNameHeader = " + xmlElement.@ColNameHeader)
                    'TextWriter.WriteLine("Format =Delimited(" + FormatDelimited & ")")
                    TextWriter.WriteLine("Format =" + xmlElement.@Format)
                    TextWriter.WriteLine("CharacterSet=" + xmlElement.@CharacterSet)
                    TextWriter.WriteLine("DecimalSymbol=" + xmlElement.@DecimalSymbol)
                    TextWriter.WriteLine("TextDelimeter = " + xmlElement.@TextDelimeter)
                    TextWriter.WriteLine("DateTimeFormat=" + xmlElement.@DateTimeFormat)

                    'generazione di tutte le colonne del FIELDS
                    Dim fields = From e In xmlN...<cols>...<col> _
                                 Select e
                    Dim iCol = 0
                    For Each f In fields

                        TextWriter.WriteLine(String.Format("Col{0}={1} {2} {3}", Threading.Interlocked.Increment(iCol).ToString, _
                                                            f.@name, f.@type, If(f.@width IsNot Nothing, " Width " + f.@width, String.Empty)))
                    Next

                    TextWriter.Flush()
                Catch s As System.Security.SecurityException
                    My.Application.Log.WriteException(s)
                    Return False
                Catch ex As Exception
                    My.Application.Log.WriteException(ex)
                    Return False
                Finally

                    If TextWriter IsNot Nothing Then TextWriter.Close()
                    If fileStream IsNot Nothing Then fileStream.Close()

                End Try

            End If

            Return True
        End Function
#End Region
    End Class

End Namespace

