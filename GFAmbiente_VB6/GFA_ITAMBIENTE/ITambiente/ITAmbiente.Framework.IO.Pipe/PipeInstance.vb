Imports System.IO.Pipes
Imports System.Threading

''' <summary>
''' 
''' </summary>
''' <remarks></remarks>
Public Class PipeInstance


    'Private _totClientsConnected As Integer
    Private _instancesServers As Hashtable
    Private _lockObj = New Object
    Private _maxInstances As Integer

    Public Event InstanceConnected As EventHandler
    Public Event InstanceDisconnected As EventHandler
    Public Event InstanceNewMessage As EventHandler(Of MessagePipe)

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Dim syncTot As New Object
    Public ReadOnly Property TotClientsConnected() As Integer
        Get
            Dim _totClientsConnected = 0

            SyncLock syncTot
                For Each e In _instancesServers.Values
                    Dim server = DirectCast(e, PipeServer)
                    If server.ConnectionPipeStates = PipeStatusEnum.PipeConnected Then
                        _totClientsConnected = _totClientsConnected + 1
                    End If
                Next

                Return _totClientsConnected

            End SyncLock
        End Get
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property InstancesServers() As Hashtable
        Get
            Return _instancesServers
        End Get
    End Property

    Public ReadOnly Property MaxInstances() As Integer
        Get
            Return _maxInstances
        End Get
    End Property



    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="name"></param>
    ''' <param name="numInstance"></param>
    ''' <remarks></remarks>
    ''' 
    <MTAThread()> _
   Public Sub Start(ByVal name As String, ByVal numInstance As Integer)


        _instancesServers = New Hashtable
        _maxInstances = numInstance
        For i = 1 To numInstance
            Dim bp = New PipeServer(name, _
                                         PipeDirection.InOut, _
                                         numInstance, _
                                         PipeTransmissionMode.Message, _
                                         PipeOptions.Asynchronous, i)

            AddHandler bp.Connected, AddressOf Connect
            AddHandler bp.Disconnected, AddressOf Disconnect
            AddHandler bp.NewMessage, AddressOf NewMessage
            InstancesServers.Add(i, bp)
            bp.Connect()
        Next

    End Sub

    'bubble event
    Private Sub NewMessage(ByVal ref As Object, ByVal m As MessagePipe)
        SyncLock _lockObj
            RaiseEvent InstanceNewMessage(ref, m)
        End SyncLock

    End Sub


    'bubble event
    Private Sub Connect(ByVal ref As Object, ByVal a As EventArgs)
        SyncLock _lockObj
            'ad ogni server corrisponde un solo client relazione 1:1
            '_totClientsConnected = _totClientsConnected + 1
#If DEBUG Then
            'System.Diagnostics.Debug.WriteLine (String.Format("Clients Connected:{0}", _totClientsConnected))
#End If
            RaiseEvent InstanceConnected(ref, a)
        End SyncLock

    End Sub

    'bubble event
    Private Sub Disconnect(ByVal ref As Object, ByVal a As PipeEventArgs)
        SyncLock _lockObj
            'If _totClientsConnected > 0 Then _totClientsConnected = _totClientsConnected - 1
            If a.Reconnect Then ref.Connect()
            RaiseEvent InstanceDisconnected(ref, a)
        End SyncLock
    End Sub



    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub [End]()
        For Each pPipe As PipeServer In InstancesServers.Values
            pPipe.Disconnect(False)
        Next
        InstancesServers.Clear()

    End Sub

    ''' <summary>
    ''' Invia un messaggio ad un client collegato se impostato clientId a blank il messaggio viene inviato a tutti i clients 
    ''' connessi
    ''' </summary>
    ''' <param name="clientId"></param>
    ''' <param name="message"></param>
    ''' <remarks></remarks>
    Public Sub SendMessage(ByVal clientId As String, ByVal message As MessagePipe)
        Dim server = From s As PipeServer In _instancesServers.Values _
                     Where (s.ClientName = clientId And clientId <> String.Empty) Or clientId = String.Empty _
                     Select s

        If server IsNot Nothing Then
            For Each sSendTo In server
                'nel caso di invio a tutti i client, viene verificato che vi sia un client connesso
                If sSendTo.ClientName IsNot Nothing Then sSendTo.SendMessagePipe(message)
            Next
        Else
            'se il client non viene trovato alzo l'eccezione
            'ToDo  traduzione
            Throw New PipeException("Client non trovato per l'invio del messaggio", Nothing)
        End If
    End Sub

End Class
