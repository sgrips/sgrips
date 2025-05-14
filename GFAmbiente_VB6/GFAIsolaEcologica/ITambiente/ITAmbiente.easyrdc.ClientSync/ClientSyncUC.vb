Imports ITAmbiente.Framework.IO.Communication.Pipe
Imports System.IO
Imports System.Runtime.CompilerServices
Imports ITAmbiente.easyrdc.Sync.PropertySettings

Public Class ClientSyncUC

    Private clientPipe As PipeClient
    Private _propertySetting As New PropertySettings
    Private _nascondi As Boolean
    Private _timer As New Timers.Timer

    Public Enum StatusMessageEnum As Integer
        NotifyConnected = 0
        NotifyDisconnected = 1
        NotifyDownload = 2
    End Enum



    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        Try

            ' Add any initialization after the InitializeComponent() call.
            PropertyGrid.SelectedObject = _propertySetting

            BWPipe.WorkerSupportsCancellation = True

            NotifyIcon.ContextMenuStrip = CtxMenuStrip

            NotifyIcon.Text = "ITAmbiente - Client Sync DB"



            My.Application.Log.DefaultFileLogWriter.CustomLocation = My.Application.Info.DirectoryPath
            My.Application.Log.DefaultFileLogWriter.BaseFileName = "Sync"
            My.Application.Log.DefaultFileLogWriter.TraceOutputOptions = TraceOptions.DateTime Or TraceOptions.ProcessId

            My.Application.Log.WriteEntry("Applicazione Avviata")


            If _propertySetting.FolderNameSync = "|CommonApplicationData|" Then
                _propertySetting.FolderNameSync = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "ITAmbiente\SyncDB")
                If Not My.Computer.FileSystem.DirectoryExists(_propertySetting.FolderNameSync) Then
                    My.Computer.FileSystem.CreateDirectory(_propertySetting.FolderNameSync)
                End If
            End If

            If _propertySetting.PathDataBase = "" Then
                _propertySetting.PathDataBase = ITAmbiente.easyrdc.Bridge.SettingsRDC.Instance.GetPropertyValue(ITAmbiente.easyrdc.Bridge.SettingsRDC.SETTING_DATABASE_PATH)
            End If

            _nascondi = _propertySetting.FolderNameSync IsNot Nothing AndAlso My.Computer.FileSystem.DirectoryExists(My.Settings.FolderNameSync) And _
                        _propertySetting.PathDataBase IsNot Nothing AndAlso My.Computer.FileSystem.DirectoryExists(My.Settings.PathDataBase) And _
                        Not String.IsNullOrEmpty(_propertySetting.ServerURL)

            gMainForm = Me

            If _propertySetting.ConnectionMode = ConnectionModeEnum.PIPE Then
                'init server Pipe
                InitServerPipe()
            Else
                'init server Shared Folder watcher
                InitWaitSharedFolder()
            End If

            If _nascondi Then
                'importo eventuali file rimasti
                ImportOldFileXML()
                'avvio il watcher sulla cartella solo se modalità folder share
                If My.Settings.ConnectionMode = ConnectionModeEnum.SHAREDFOLDER Then

                    AddHandler FileSystemWatcher.Renamed, AddressOf RenameFileEvent
                    AddHandler _timer.Elapsed, AddressOf SendMovimentiToServer
                    '
                    InitWatcher()
                    '
                    _timer.Interval = 1000 * 60
                    _timer.AutoReset = True
                    _timer.Enabled = False

                    _timer.Start()
                End If


                Me.WindowState = FormWindowState.Minimized

            End If

        Catch ex As Exception
            My.Application.Log.WriteException(ex)
        End Try

        ChangeIconStatus(StatusMessageEnum.NotifyDisconnected)

    End Sub


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitWatcher()
        FileSystemWatcher.Path = _propertySetting.FolderNameSync
        FileSystemWatcher.NotifyFilter = IO.NotifyFilters.FileName Or IO.NotifyFilters.CreationTime Or IO.NotifyFilters.Size
        FileSystemWatcher.Filter = _propertySetting.FileNamePattern
        FileSystemWatcher.EnableRaisingEvents = True
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        _propertySetting.Save()
        InitWatcher()
        Me.WindowState = FormWindowState.Minimized

        'chiudo server pipe
        EndServerPipe()
        EndSharedFolderWatcher()

        If _propertySetting.ConnectionMode = ConnectionModeEnum.PIPE Then
            'init server Pipe
            InitServerPipe()
        Else
            InitWaitSharedFolder()
        End If

    End Sub

    Private Sub NotifyIcon_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon.MouseClick
        If e.Button = Windows.Forms.MouseButtons.Left Then NotifyIcon.ShowBalloonTip(500)
    End Sub



    Private Sub NotifyIcon_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon.MouseDoubleClick
        ImpostazioniToolStripMenuItem.PerformClick()

    End Sub

    ''' <summary>
    ''' Modifica l'icona di stato 
    ''' </summary>
    ''' <param name="State"></param>
    ''' <remarks></remarks>
    Public Sub ChangeIconStatus(ByVal State As StatusMessageEnum)

        If Me.NotifyIcon.Icon IsNot Nothing Then Me.NotifyIcon.Icon.Dispose()


        If State = StatusMessageEnum.NotifyDownload Then
            ' Me.NotifyIcon.Icon = Nothing
        Else
            State = If(My.Computer.Network.IsAvailable And State = StatusMessageEnum.NotifyConnected, StatusMessageEnum.NotifyConnected, StatusMessageEnum.NotifyDisconnected)
            Me.NotifyIcon.Icon = If(State = StatusMessageEnum.NotifyConnected, My.Resources.connect, My.Resources.pending)
        End If

        If clientPipe IsNot Nothing AndAlso clientPipe.ConnectionPipeStates = PipeStatusEnum.PipeConnected Then
            clientPipe.SendMessagePipe(New MessagePipe(MessagePipeIDEnum.MessagePipeClientStatus, State))
        End If


    End Sub

    Private Sub ClientSyncUC_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        NotifyIcon.Dispose()
        _timer.Stop()

        'chiudo server pipe
        EndServerPipe()

        'chiudo shared folder watcher
        EndSharedFolderWatcher()

        My.Application.Log.WriteEntry("Applicazione Fermata")

        Util.Connection.CloseMDB()


    End Sub


    Private Sub Setting_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
        If Me.WindowState = FormWindowState.Minimized Then
            _nascondi = True
            Me.Hide()
        End If
    End Sub


    Private Sub Setting_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
        If _nascondi Then Me.Hide()
    End Sub

    Private Sub ImpostazioniToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImpostazioniToolStripMenuItem.Click
        _nascondi = False
        Application.DoEvents()
        Me.Visible = True
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.WindowState = FormWindowState.Normal
        FileSystemWatcher.EnableRaisingEvents = False
    End Sub

    Private Sub mnuEsci_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEsci.Click
        Me.Close()
    End Sub

    ''' <summary>
    ''' Dichiarazione del Delegato per permettere la chiusura della FORM 
    ''' </summary>
    ''' <remarks></remarks>
    Delegate Sub DelegateCloseForm()

    ''' <summary>
    ''' Metodo per il delegato DelegateCloseForm
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CloseForm()
        Me.Close()
    End Sub

#Region "Shared Folder Watcher"
    Private ServerToPing As Dictionary(Of String, SharedFolderListen) = New Dictionary(Of String, SharedFolderListen)
    ''' <summary>
    ''' avvio il processo background per il monitoraggio della cartella condivisa sul server
    ''' </summary>
    ''' <remarks></remarks>
    Sub InitWaitSharedFolder()
        'creo una classe di tipo 
        Dim folderLister As SharedFolderListen = New SharedFolderListen(My.Settings.SharedFolderName, My.Settings.ServerURL)
        AddHandler folderLister.ServerConnected, AddressOf ConnectedToPC
        ServerToPing.Add(My.Settings.SharedFolderName, folderLister)

    End Sub

    Dim methodSync As Object = New Object
    Sub ConnectedToPC(ByVal sender As Object, ByVal e As EventArgs)
        SyncLock methodSync

            Dim obj = DirectCast(sender, SharedFolderListen)

            RefreshNotifyClientsStatus()

        End SyncLock

    End Sub

    Sub RefreshNotifyClientsStatus()
        Dim sb = New System.Text.StringBuilder("")

        sb.Append("Sincronizzazione dati IsolaEcologica." & vbCrLf & _
                                    "Stato Connessioni" & vbCrLf)

        For Each c In ServerToPing.Values
            sb.Append(String.Format("Unita Server {0} - {1}" & vbCrLf, c.PCName, _
                                    If(c.State = SharedFolderListen.SharedFolderStateEnum.Connected, "Connesso", "Non Connesso")))

            ChangeIconStatus(If(c.State = SharedFolderListen.SharedFolderStateEnum.Connected, StatusMessageEnum.NotifyConnected, StatusMessageEnum.NotifyDisconnected))
           
        Next

        NotifyIcon.BalloonTipIcon = ToolTipIcon.Info
        NotifyIcon.BalloonTipTitle = "Client Sync"
        NotifyIcon.BalloonTipText = sb.ToString

    End Sub



    ''' <summary>
    ''' chiusura server watcher
    ''' </summary>
    ''' <remarks></remarks>
    <MethodImplAttribute(MethodImplOptions.Synchronized)> _
    Sub EndSharedFolderWatcher()
        For Each cp In ServerToPing
            cp.Value.Close()
        Next
        ServerToPing.Clear()

        My.Application.Log.WriteEntry("Client Shared Folder Stop", TraceEventType.Verbose)
    End Sub
#End Region

#Region "PIPE"

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BckGndWrkClietPipe_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BWPipe.DoWork
        Dim connected As Boolean = False

        Dim bw As System.ComponentModel.BackgroundWorker = CType(sender, System.ComponentModel.BackgroundWorker)

        Do While Not bw.CancellationPending And Not connected
            Try
                clientPipe.Connect()
                connected = True
            Catch ex As PipeException
                'avvio processo in background per attesa connessione con server
                Threading.Thread.Sleep(1000)
            End Try
        Loop

        'chiusura istanze server pipe
        clientPipe.Disconnect()
    End Sub

    ''' <summary>
    ''' Inizializzazione server pipe
    ''' </summary>
    ''' <remarks></remarks>
    Sub InitServerPipe()
        'attiva il clientPipe per la ricezione e l'invio di messaggi al serverPipe
        clientPipe = New PipeClient(_propertySetting.ServerURL, "ITAmbienteSync", _propertySetting.IDClient, System.IO.Pipes.PipeDirection.InOut)


        AddHandler clientPipe.Connected, AddressOf ClientConnected
        AddHandler clientPipe.Disconnected, AddressOf ClientDisconnected
        AddHandler clientPipe.NewMessage, AddressOf ClientNewMessage

        BWPipe.WorkerSupportsCancellation = True
        BWPipe.RunWorkerAsync()
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClientConnected(ByVal sender As Object, ByVal e As EventArgs)
        ChangeIconStatus(StatusMessageEnum.NotifyConnected)
    End Sub



    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClientDisconnected(ByVal sender As Object, ByVal e As EventArgs)
        'se il client viene disconnesso allora ritento la connessione con il server
        RemoveHandler clientPipe.Connected, AddressOf ClientConnected
        RemoveHandler clientPipe.Disconnected, AddressOf ClientDisconnected
        RemoveHandler clientPipe.NewMessage, AddressOf ClientNewMessage

        clientPipe = Nothing

        If Not BWPipe.CancellationPending Then BWPipe.RunWorkerAsync()

        ChangeIconStatus(StatusMessageEnum.NotifyDisconnected)

    End Sub

    ''' <summary>
    ''' chiusura server pipe
    ''' </summary>
    ''' <remarks></remarks>
    <MethodImplAttribute(MethodImplOptions.Synchronized)> _
    Sub EndServerPipe()

        If clientPipe Is Nothing Then Exit Sub

        RemoveHandler clientPipe.Connected, AddressOf ClientConnected
        RemoveHandler clientPipe.Disconnected, AddressOf ClientDisconnected
        RemoveHandler clientPipe.NewMessage, AddressOf ClientNewMessage

        BWPipe.CancelAsync()

        Do While BWPipe.IsBusy
            Application.DoEvents()
        Loop

        My.Application.Log.WriteEntry("Server Client Stop", TraceEventType.Verbose)
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="msg"></param>
    ''' <param name="idMsg"></param>
    ''' <param name="idCode"></param>
    ''' <remarks></remarks>
    Public Sub SendMessagePipe(ByVal msg As String, ByVal idMsg As MessagePipeIDEnum, ByVal idCode As Integer)
        If clientPipe.ConnectionPipeStates = PipeStatusEnum.PipeConnected Then clientPipe.SendMessagePipe(New MessagePipe(idMsg, idCode, msg))
    End Sub



    ''' <summary>
    ''' Messaggio ricevuto dal SERVER PIPE
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClientNewMessage(ByVal sender As Object, ByVal m As MessagePipe)
        Select Case m.Id
            Case MessagePipeIDEnum.MessagePipeClientEnd
                'Dim closeForm As New DelegateCloseForm(AddressOf Me.CloseForm)
                'Me.Invoke(closeForm)
                BWPipe.CancelAsync()

                Do While BWPipe.IsBusy
                    Application.DoEvents()
                Loop

                Threading.Thread.Sleep(500)

                clientPipe.Disconnect()


            Case MessagePipeIDEnum.MessagePipeClientRefresh
                'refresh del servizio

            Case MessagePipeIDEnum.MessagePipeClientInfo
                Select Case m.MessageCode
                    Case Pipes.Messagecode.Constants.MESSAGE_CODE_FILESYNC
                        StartImport(Path.Combine(My.Settings.FolderNameSync, m.MessageText))
                    Case Pipes.Messagecode.Constants.MESSAGE_CODE_IMPORTOLD
                        ImportOldFileXML()
                    Case Pipes.Messagecode.Constants.MESSAGE_CODE_GETMOVIMENTI
                        SendMovimentiToServer()
                End Select

        End Select
    End Sub

#End Region
End Class
