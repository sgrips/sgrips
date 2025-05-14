Public Class EntityID
    Dim _cmdNumeratoriINS As New OleDb.OleDbCommand With {.CommandType = CommandType.Text}
    Dim _cmdNumeratoriUPD As New OleDb.OleDbCommand With {.CommandType = CommandType.Text}
    Dim _cmdNumeratori As New OleDb.OleDbCommand With {.CommandType = CommandType.Text}
    Dim _entityID As New Hashtable


    Public Sub New()
        Dim p = New OleDb.OleDbParameter("Entita", OleDb.OleDbType.VarChar)
        _cmdNumeratori.Parameters.Add(p)
        _cmdNumeratori.CommandText = "Select Numeratore from Contatori where Entita=?"

        p = New OleDb.OleDbParameter("Numeratore", OleDb.OleDbType.Integer)
        _cmdNumeratoriUPD.Parameters.Add(p)
        p = New OleDb.OleDbParameter("Entita", OleDb.OleDbType.VarChar)
        _cmdNumeratoriUPD.Parameters.Add(p)
        _cmdNumeratoriUPD.CommandText = "update Contatori set  Numeratore=? where Entita=?"


        p = New OleDb.OleDbParameter("Numeratore", OleDb.OleDbType.Integer)
        _cmdNumeratoriINS.Parameters.Add(p)
        p = New OleDb.OleDbParameter("Entita", OleDb.OleDbType.VarChar)
        _cmdNumeratoriINS.Parameters.Add(p)
        _cmdNumeratoriINS.CommandText = "insert into Contatori (Numeratore,Entita) values(? ,?)"
    End Sub

    Public Sub Close()
        _cmdNumeratori.Dispose()
        _cmdNumeratoriINS.Dispose()
        _cmdNumeratoriUPD.Dispose()
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="entityID"></param>
    ''' <param name="oleDBtrans"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNewEntityID(ByVal entityID As String, Optional ByVal oleDBtrans As OleDb.OleDbTransaction = Nothing)
        Dim item As Integer

        If Not _entityID.ContainsKey(entityID) Then
            item = GetNewID(entityID, oleDBtrans)
            _entityID.Add(entityID, item)
        Else
            item = DirectCast(_entityID.Item(entityID), Integer)
            item = item + 1
            _entityID.Item(entityID) = item
        End If

        Return item
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="oleDBtrans"></param>
    ''' <remarks></remarks>
    Public Sub SaveAll(ByVal oleDBtrans As OleDb.OleDbTransaction)
        Try
            For Each e In _entityID.Keys

                Dim value = _entityID.Item(e)
                Dim key = e
                If GetNewID(key, oleDBtrans) = 1 Then
                    _cmdNumeratoriINS.Parameters("Entita").Value = key
                    _cmdNumeratoriINS.Parameters("Numeratore").Value = value
                    _cmdNumeratoriINS.Transaction = oleDBtrans
                    _cmdNumeratoriINS.Connection = Util.Connection.ConnectionMDB
                    _cmdNumeratoriINS.ExecuteNonQuery()

                Else
                    _cmdNumeratoriUPD.Parameters("Entita").Value = key
                    _cmdNumeratoriUPD.Parameters("Numeratore").Value = value
                    _cmdNumeratoriUPD.Transaction = oleDBtrans
                    _cmdNumeratoriUPD.Connection = Util.Connection.ConnectionMDB
                    _cmdNumeratoriUPD.ExecuteNonQuery()
                End If
            Next

        Catch ex As Exception
            My.Application.Log.WriteException(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="ContType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetNewID(ByVal ContType As String, ByVal oleDBtrans As OleDb.OleDbTransaction) As Integer
        _cmdNumeratori.Parameters(0).Value = ContType
        Try
            _cmdNumeratori.Connection = Util.Connection.ConnectionMDB
            _cmdNumeratori.Transaction = oleDBtrans


            Dim value = _cmdNumeratori.ExecuteScalar
            value = If(value Is Nothing, 1, value + 1)

            Return value

        Catch ex As Exception
            Return -1
        End Try

    End Function
End Class
