Imports System.Threading

Public Class SharedFolderListen

    Public Delegate Sub EventHandlerServerConnected(ByVal sender As Object, ByVal e As EventArgs)
    Public Event ServerConnected(ByVal sender As Object, ByVal e As EventArgs)

    Private _sharedFolderName As String
    Public Property SharedFolderName() As String
        Get
            Return _sharedFolderName
        End Get
        Set(ByVal value As String)
            _sharedFolderName = value
        End Set
    End Property

    Private _PCName As String
    Public Property PCName() As String
        Get
            Return _PCName
        End Get
        Set(ByVal value As String)
            _PCName = value
        End Set
    End Property

    Private _idPesa As String
    Public Property IDPesa() As String
        Get
            Return _idPesa
        End Get
        Set(ByVal value As String)
            _idPesa = value
        End Set
    End Property


    Private _pingthread As Threading.Thread
    Public Property PingThread() As Threading.Thread
        Get
            Return _pingthread
        End Get
        Set(ByVal value As Threading.Thread)
            _pingthread = value
        End Set
    End Property

    Enum SharedFolderStateEnum
        Connected
        Disconnected
    End Enum

    Private _state As SharedFolderStateEnum
    Public Property State() As SharedFolderStateEnum
        Get
            Return _state
        End Get
        Set(ByVal value As SharedFolderStateEnum)
            _state = value
        End Set
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sharedFolder"></param>
    ''' <param name="serverName"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal sharedFolder As String, ByVal serverName As String)
        Me.PCName = serverName
        Me.SharedFolderName = sharedFolder
        Me.State = SharedFolderStateEnum.Disconnected
        PingThread = New Threading.Thread(AddressOf StartPing)
        PingThread.Start()

    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sharedFolder"></param>
    ''' <param name="serverName"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal sharedFolder As String, ByVal serverName As String, ByVal idPesa As String)
        Me.New(sharedFolder, serverName)
        Me.IDPesa = idPesa
    End Sub


    ''' <summary>
    ''' faccio il ping al server per verificare quanto attivo 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub StartPing()
        Try
            Do While True
                Try
                    If My.Computer.Network.Ping(PCName, 1000) Then
                        Me.State = SharedFolderStateEnum.Connected
                        RaiseEvent ServerConnected(Me, Nothing)
                        'se connesso i test li faccio ogni 30 sec
                        Threading.Thread.Sleep(1000 * 30)
                    Else
                        Threading.Thread.Sleep(1000)
                        Me.State = SharedFolderStateEnum.Disconnected
                    End If

                Catch ex As Net.NetworkInformation.PingException
                    '    Console.WriteLine(String.Format("Client {0} non trovato", PCName))
                End Try
            Loop

        Catch ex As ThreadAbortException
            Console.WriteLine( _
                CType(ex.ExceptionState, String))
        Catch ex As Exception
            My.Application.Log.WriteException(ex)

        End Try

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Close()
        If PingThread IsNot Nothing AndAlso PingThread.IsAlive Then
            PingThread.Abort("Terminate Work")
            PingThread.Join()
        End If
    End Sub


End Class
