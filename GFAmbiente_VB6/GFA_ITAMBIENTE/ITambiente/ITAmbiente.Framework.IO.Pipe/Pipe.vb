Imports System.IO.Pipes
Imports System.Threading
Imports System.IO
Imports System.Xml.Serialization

''' <summary>
''' 
''' </summary>
''' <remarks></remarks>
Public Class MessagePipe
    Inherits EventArgs

    Private _messageByte As Byte()
    Private _id As MessagePipeIDEnum
    Private Shared _lock As Object = New Object


    Private _messageCode As Integer
    Private Const MESSAGE_CODE_NONE As Integer = -1



    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
    End Sub


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="message"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal message As String)
        Me.New(MessagePipeIDEnum.MessagePipeClientInfo, MESSAGE_CODE_NONE, message)
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="id">message id</param>
    ''' <param name="message">text message</param>
    ''' <remarks></remarks>
    Sub New(ByVal id As MessagePipeIDEnum, ByVal message As String)
        Me.New(id, MESSAGE_CODE_NONE, message)
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="id">message id</param>
    ''' <param name="messageCode">custom code,used for indicate more message detail</param> 
    ''' <param name="message">text message</param>
    ''' <remarks></remarks>
    Sub New(ByVal id As MessagePipeIDEnum, ByVal messageCode As Integer, ByVal message As String)
        _id = id
        MessageText = message
        _messageCode = messageCode
    End Sub


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlIgnoreAttribute()> _
    Property MessageText() As String
        Get
            Return Text.UnicodeEncoding.Unicode.GetString(_messageByte)
        End Get
        Set(ByVal value As String)
            _messageByte = Text.UnicodeEncoding.Unicode.GetBytes(value)
        End Set
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlElement(elementname:="base64Binary", Type:=GetType(Byte()))> _
    Property MessageByte() As Byte()
        Get
            Return _messageByte
        End Get
        Set(ByVal value As Byte())
            _messageByte = value
        End Set
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Id() As MessagePipeIDEnum
        Get
            Return _id
        End Get
        Set(ByVal value As MessagePipeIDEnum)
            _id = value
        End Set
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property MessageCode() As Integer
        Get
            Return _messageCode
        End Get
        Set(ByVal value As Integer)
            _messageCode = value
        End Set
    End Property
End Class

#Region "Server Pipe"
''' <summary>
''' Classe che rappresenta una Pipe Server
''' </summary>
Friend Class PipeServer
    Inherits Pipe


    Private _maxInstances As Integer
    Private _namedPipeServerStream As NamedPipeServerStream
    Private _iAsyncResult As IAsyncResult
    Private _lockConnect As New Object
    Private _lockDisconnect As New Object
    Private _instanceNum As Integer
    Private _clientName As String


    Public Event Connected As EventHandler
    Shadows Event Disconnected As EventHandler



    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="name"></param>
    ''' <param name="direction"></param>
    ''' <param name="maxInstances"></param>
    ''' <param name="pipeTransmissionMode"></param>
    ''' <param name="pipeOptions"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal name As String, _
        ByVal direction As PipeDirection, _
        ByVal maxInstances As Integer, _
        ByVal pipeTransmissionMode As System.IO.Pipes.PipeTransmissionMode, _
        ByVal pipeOptions As PipeOptions, _
        ByVal instanceNum As Integer)


        Me.Name = name
        Me.Direction = direction
        Me.PipeTransmissionMode = pipeTransmissionMode
        Me.PipeOptions = pipeOptions
        _instanceNum = instanceNum

        _maxInstances = maxInstances

        If String.IsNullOrEmpty(Me.Name) Then
            Throw New ArgumentException("name parameter not valid")
        End If

        If String.IsNullOrEmpty(Me.Direction) Then
            Throw New ArgumentException("direction parameter not valid")
        End If

        If String.IsNullOrEmpty(Me.MaxInstances) Then
            Throw New ArgumentException("maxInstances parameter not valid")
        End If

        If String.IsNullOrEmpty(Me.PipeTransmissionMode) Then
            Throw New ArgumentException("pipeTransmissionMode parameter not valid")
        End If

        If String.IsNullOrEmpty(Me.PipeOptions) Then
            Throw New ArgumentException("pipeOptions parameter not valid")
        End If

        Try
            _namedPipeServerStream = New NamedPipeServerStream(Me.Name, Me.Direction, Me.MaxInstances, Me.PipeTransmissionMode, Me.PipeOptions)
        Catch ex As Exception
            Throw New PipeException(ex.Message, ex)
        End Try


        Me.PipeStream = _namedPipeServerStream

    End Sub


    Protected Overrides Sub BeforeRaiseEventNewMessage(ByVal sender As Object, ByVal m As MessagePipe)
        Select Case m.Id
            Case MessagePipeIDEnum.MessagePipeClientName
                'associo al Server il client impostando la property ClientName 
                _clientName = m.MessageText
            Case MessagePipeIDEnum.MessagePipeServerEnd
                'il client si disconnette quindi chiudo anche la relativa referenza server
                Me.Disconnect()
        End Select
    End Sub

    ''' <summary>
    ''' Numero Massimo di istanze creabile del server
    ''' </summary>
    Public Property MaxInstances() As Integer
        Get
            Return _maxInstances
        End Get
        Set(ByVal value As Integer)
            _maxInstances = value
        End Set
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ReadOnly Property InstanceNum()
        Get
            Return _instanceNum
        End Get
    End Property


    ''' <summary>
    ''' Chiude la attesa di connesione da parte di un client
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub CancelRequestConnection()
        If _iAsyncResult IsNot Nothing Then
            Try
                _namedPipeServerStream.EndWaitForConnection(_iAsyncResult)
            Catch ex As Exception
            End Try
            _iAsyncResult = Nothing
        End If
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Disconnect(Optional ByVal reconnect As Boolean = True)
        Try
            If _namedPipeServerStream.IsConnected Then
                'termina la lettura

                _namedPipeServerStream.Disconnect()
            End If

        Catch ex As Exception
        Finally
            Try
                Me.CancelRequestConnection()
                _namedPipeServerStream.Close()
                _namedPipeServerStream.Dispose()
            Catch ex As Exception

            End Try

            System.Diagnostics.Debug.WriteLine(String.Format("Server {0} Disconnected con ID {1}", Me.InstanceNum, Me.Name))

            Me.ConnectionPipeStates = PipeStatusEnum.PipeDisconnected

            RaiseEvent Disconnected(Me, New PipeEventArgs(reconnect, _clientName))

        End Try
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Connect()
        Try
            _iAsyncResult = _namedPipeServerStream.BeginWaitForConnection(AddressOf ClientConnected, Nothing)
        Catch ex As Exception
            Throw New PipeException(ex.Message, ex)
        End Try
    End Sub

    ''' <summary>
    ''' Connection Waitting
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClientConnected(ByVal obj As Object)

        Try
            Me.CancelRequestConnection()
            Me.ConnectionPipeStates = PipeStatusEnum.PipeConnected

            _namedPipeServerStream.ReadMode = Pipes.PipeTransmissionMode.Message

            RaiseEvent Connected(Me, Nothing)

            Me.ReadBegin()

        Catch ex As Exception
            Throw New PipeException(ex.Message, ex)
        End Try

    End Sub

    Private Sub BartPipeServer_NewMessage(ByVal sender As Object, ByVal e As MessagePipe) Handles Me.NewMessage
        System.Diagnostics.Debug.WriteLine(String.Format("Client {0} Wrote: {1}", _clientName, e.MessageText))
#If DEBUG Then
        Me.SendMessagePipe(New MessagePipe("Ok da server:" + InstanceNum.ToString))
#End If
    End Sub

    ''' <summary>
    ''' Nome del Client collegato
    ''' </summary>
    Public Property ClientName() As String
        Get
            Return _clientName
        End Get
        Set(ByVal value As String)
            _clientName = value
        End Set
    End Property
End Class
#End Region


#Region "Client"
Public Class PipeClient
    Inherits Pipe


    Private _namedPipeClientStream As NamedPipeClientStream
    Private _lockConnect As New Object
    Private _lockDisconnect As New Object
    Private _iAsyncResult As IAsyncResult
    Private _serverName As String

    Public Event Connected As EventHandler
    Shadows Event Disconnected As EventHandler

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="serverURL"></param>
    ''' <param name="serverName"></param>
    ''' <param name="direction"></param>
    ''' <remarks></remarks>
    Sub New(ByVal serverURL As String, ByVal serverName As String, ByVal name As String, ByVal direction As System.IO.Pipes.PipeDirection)


        _serverName = serverName
        Me.Name = name
        Me.Direction = direction

        If String.IsNullOrEmpty(Me.Name) Then
            Throw New ArgumentException("name parameter not valid")
        End If

        If String.IsNullOrEmpty(Me.Direction) Then
            Throw New ArgumentException("direction parameter not valid")
        End If

        If String.IsNullOrEmpty(serverName) Then
            Throw New ArgumentException("pipeTransmissionMode parameter not valid")
        End If

        Try
            _namedPipeClientStream = New NamedPipeClientStream(serverURL, serverName, direction, PipeOptions.Asynchronous, _
                                                       Security.Principal.TokenImpersonationLevel.Impersonation)

        Catch ex As Exception
            Throw New PipeException(ex.Message, ex)
        End Try

        Me.PipeStream = _namedPipeClientStream
    End Sub



    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ReadOnly Property ServerName()
        Get
            Return _serverName
        End Get

    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Disconnect(Optional ByVal reconnect As Boolean = True)
        Try
            'invio messaggio al serve di notifica chiusura
            Try
                SendMessagePipe(New MessagePipe() With {.Id = MessagePipeIDEnum.MessagePipeServerEnd, .MessageText = 1})
            Catch ex As Exception

            End Try


            If _namedPipeClientStream.IsConnected Then _namedPipeClientStream.Close()
        Catch ex As Exception
        Finally
            Try
                _namedPipeClientStream.Close()
                _namedPipeClientStream.Dispose()
            Catch ex As Exception

            End Try
            System.Diagnostics.Debug.WriteLine("Client Disconnected")

            Me.ConnectionPipeStates = PipeStatusEnum.PipeDisconnected
            RaiseEvent Disconnected(Me, Nothing)
        End Try
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="m"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub BeforeRaiseEventNewMessage(ByVal sender As Object, ByVal m As MessagePipe)

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Connect()

        Try
            If _namedPipeClientStream.IsConnected = False Then

                System.Diagnostics.Debug.WriteLine(String.Format("Attempting to connect to the pipe {0}", Me.ServerName))

                'sincrona
                Me.ConnectionPipeStates = PipeStatusEnum.PipeWaitConnecting

                _namedPipeClientStream.Connect(10)

                System.Diagnostics.Debug.WriteLine("Connected to the pipe ok")

                _namedPipeClientStream.ReadMode = PipeTransmissionMode.Message
                'inizia lettura asincrona delle stream
                Me.ReadBegin()

                Me.ConnectionPipeStates = PipeStatusEnum.PipeConnected

                'invia al server il nome del client pipe
                SendMessagePipe(New MessagePipe() With {.Id = MessagePipeIDEnum.MessagePipeClientName, .MessageText = Me.Name})

                RaiseEvent Connected(Me, Nothing)


            End If


        Catch ex As Exception
            Throw New PipeException(ex.Message, ex)
        End Try
    End Sub

    Private Sub BartPipeClient_NewMessage(ByVal sender As Object, ByVal e As MessagePipe) Handles Me.NewMessage
        System.Diagnostics.Debug.WriteLine(String.Format("Received from server: {0}", e.MessageText))
    End Sub
End Class
#End Region


#Region "BartPipe"
''' <summary>
''' 
''' </summary>
''' <remarks></remarks>
Public MustInherit Class Pipe

    Private _connectionPipeStates As PipeStatusEnum = PipeStatusEnum.PipeDisconnected
    Private _direction As PipeDirection
    Private _pipeTransmissionMode As System.IO.Pipes.PipeTransmissionMode
    Private _pipeOptions As PipeOptions
    Private _name As String
    Private _pipeStream As PipeStream

    'events
    Public Event NewMessage As EventHandler(Of MessagePipe)

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PipeStream() As PipeStream
        Get
            Return _pipeStream
        End Get
        Set(ByVal value As PipeStream)
            _pipeStream = value

            Try
                '_sw = New StreamWriter(PipeStream)
                '_sr = New StreamReader(PipeStream)
            Catch ex As Exception
                Throw New PipeException(ex.Message, ex)
            End Try
        End Set
    End Property

    ''' <summary>
    ''' Thread corrente
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property CurrentThread() As Thread
        Get
            Return Thread.CurrentThread
        End Get
    End Property

    ''' <summary>
    ''' Direzione delle Pipe
    ''' </summary>
    ''' <value>PipeDirection.In , PipeDirection.Out o PipeDirection.InOut</value>
    Public Property Direction() As PipeDirection
        Get
            Return _direction
        End Get
        Set(ByVal value As PipeDirection)
            _direction = value
        End Set
    End Property

    ''' <summary>
    ''' Opzioni
    ''' </summary>
    Public Property PipeOptions() As PipeOptions
        Get
            Return _pipeOptions
        End Get
        Set(ByVal value As PipeOptions)
            _pipeOptions = value
        End Set
    End Property

    ''' <summary>
    ''' Tipo di trasmissione dati
    ''' </summary>
    Public Property PipeTransmissionMode() As PipeTransmissionMode
        Get
            Return _pipeTransmissionMode
        End Get
        Set(ByVal value As PipeTransmissionMode)
            _pipeTransmissionMode = value
        End Set
    End Property

    ''' <summary>
    ''' Nome della Pipe
    ''' </summary>
    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
        End Set
    End Property

    ''' <summary>
    ''' Stato della connesione 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ConnectionPipeStates() As PipeStatusEnum
        Get
            Dim c As PipeStream = Me.PipeStream

            Return If(_connectionPipeStates = PipeStatusEnum.PipeConnected And c.IsConnected, PipeStatusEnum.PipeConnected, PipeStatusEnum.PipeDisconnected)
        End Get
        Set(ByVal value As PipeStatusEnum)
            _connectionPipeStates = value
        End Set
    End Property


    Class StateRead
        Dim _buffer As New Text.StringBuilder
        Dim _bufferArray() As Byte

        Public Sub New()
        End Sub

        WriteOnly Property BufferArray() As Byte()
            Set(ByVal value As Byte())
                _bufferArray = value
            End Set
        End Property

        ReadOnly Property BufferToString() As String
            Get
                Return _buffer.ToString
            End Get
        End Property

        Sub AddBuffer(ByVal numByte As Integer)
            _buffer.Append(Text.UTF8Encoding.UTF8.GetChars(_bufferArray, 0, numByte))
        End Sub
    End Class

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub ReadBegin()
        Dim arrayByte(1024) As Byte
        Dim st = New StateRead

        Try
            st.BufferArray = arrayByte
            _pipeStream.BeginRead(arrayByte, 0, arrayByte.Length, AddressOf ReadSomething, st)

        Catch ex As ObjectDisposedException
            'la pipe è chiusa

        Catch ex As InvalidOperationException
            If Not _pipeStream.IsConnected And _pipeStream.GetType Is GetType(PipeClient) Then
                Me.Disconnect()
            End If

        Catch ex As Exception

        End Try
    End Sub


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub ReadSomething(ByVal _iAsyncResult As IAsyncResult)
        Dim numByte As Integer
        Dim st = DirectCast(_iAsyncResult.AsyncState, StateRead)
        Try

            If Me._pipeStream.IsConnected Then
                numByte = _pipeStream.EndRead(_iAsyncResult)
                st.AddBuffer(numByte)
                If _pipeStream.IsMessageComplete Then
                    If numByte > 0 Then
                        If _pipeStream.IsMessageComplete Then
                            Dim mp As MessagePipe
                            Dim sr = New StringReader(st.BufferToString)

                            ' Construct an instance of the XmlSerializer with the type
                            ' of object that is being deserialized.
                            Dim mySerializer As XmlSerializer = New XmlSerializer(GetType(MessagePipe))
                            ' Call the Deserialize method and cast to the object type.
                            mp = CType(mySerializer.Deserialize(sr), MessagePipe)

                            BeforeRaiseEventNewMessage(Me, mp)

                            RaiseEvent NewMessage(Me, mp)
                        End If
                    End If

                    If Me._pipeStream.IsConnected And numByte > 0 Then
                        ReadBegin()
                    End If
                Else
                    Dim arrayByte(1024) As Byte
                    st.BufferArray = arrayByte
                    _pipeStream.BeginRead(arrayByte, 0, arrayByte.Length, AddressOf ReadSomething, st)
                End If


            End If
        Catch ex As Exception
            'TODO decidere cosa fare in caso di errore
            System.Diagnostics.Debug.WriteLine(ex)
            Console.WriteLine(ex)
            ReadBegin()
        End Try
    End Sub


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CallBackWriteAsync(ByVal result As IAsyncResult)
        Try

            If result.IsCompleted Then
                _pipeStream.EndWrite(result)
                Diagnostics.Debug.WriteLine("End per:" + result.GetHashCode.ToString)
            End If
        Catch ex As Exception

        End Try
    End Sub


    Dim objSync As New Object
    ''' <summary>
    ''' Send a message type messagePipe to server
    ''' </summary>
    ''' <param name="message"></param>
    ''' <remarks></remarks>
    Public Sub SendMessagePipe(ByVal message As MessagePipe)

        SyncLock objSync
            Dim mySerializer As XmlSerializer = New XmlSerializer(GetType(MessagePipe))

            Try
                If (TypeOf (Me) Is PipeClient AndAlso _
                    Me.ConnectionPipeStates = PipeStatusEnum.PipeConnected) Or _
                    TypeOf (Me) Is PipeServer Then


                    Dim m = New MemoryStream

                    mySerializer.Serialize(m, message)

                    Dim byteAtt As Byte() = m.ToArray



                    Dim _asyncResultWrite = _pipeStream.BeginWrite(byteAtt, 0, byteAtt.Length, AddressOf CallBackWriteAsync, Nothing)

                    Diagnostics.Debug.WriteLine("Attesa per:" + _asyncResultWrite.GetHashCode.ToString)


                End If

            Catch ex As IOException

            Catch ex As Exception
                Throw New PipeException(ex.Message, ex)
            End Try


        End SyncLock



    End Sub



    Public MustOverride Sub Connect()
    Public MustOverride Sub Disconnect(Optional ByVal reconnect As Boolean = True)

    Protected MustOverride Sub BeforeRaiseEventNewMessage(ByVal sender As Object, ByVal m As MessagePipe)

    Public Sub New()

    End Sub
End Class
#End Region

''' <summary>
''' 
''' </summary>
''' <remarks></remarks>
Public Class PipeException
    Inherits System.ApplicationException

    Public Sub New(ByVal message As String, ByVal innerException As Exception)
        MyBase.New(message, innerException)
    End Sub

End Class

''' <summary>
''' 
''' </summary>
''' <remarks></remarks>
Friend Class PipeEventArgs
    Inherits EventArgs

    Dim _reconnect
    Dim _clientName

    Sub New(ByVal reconnect As Boolean, ByVal clientName As String)
        _reconnect = reconnect
    End Sub

    ReadOnly Property Reconnect() As Boolean
        Get
            Return _reconnect
        End Get
    End Property

    ReadOnly Property ClientName() As String
        Get
            Return _clientName
        End Get
    End Property
End Class
