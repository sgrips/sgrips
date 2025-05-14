Imports System.ComponentModel
Imports System.Timers
Imports Microsoft.Win32
Imports System.IO
Imports ITAmbiente.Framework.IO.Communication.Pipe
Imports System.Runtime.CompilerServices

Public Class Setting
    Private _nascondi As Boolean = True
    Private _propertySetting As New PropertySettings
    Private _sync As SyncToDB
    Private _timer As New Timers.Timer
    Private _listOfClients As New Collections.ArrayList

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        _nascondi = False

        ' Add any initialization after the InitializeComponent() call.
        PropertyGrid.SelectedObject = _propertySetting

        NotifyIcon.ContextMenuStrip = CtxMenuStrip

        NotifyIcon.Text = "ITAmbiente - Server Sync DB"

        My.Application.Log.DefaultFileLogWriter.CustomLocation = My.Application.Info.DirectoryPath
        My.Application.Log.DefaultFileLogWriter.BaseFileName = "Sync"
        My.Application.Log.DefaultFileLogWriter.TraceOutputOptions = TraceOptions.DateTime Or TraceOptions.ProcessId

        My.Application.Log.WriteEntry("Applicazione Avviata")



        _sync = New SyncToDB
        _sync.MainForm = Me
        Try

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
                        _propertySetting.PathDataBase IsNot Nothing AndAlso My.Computer.FileSystem.DirectoryExists(My.Settings.PathDataBase)

            AddHandler FileSystemWatcher.Renamed, AddressOf _sync.RenameFileEvent


            _timer.Interval = 1000 * 60
            _timer.AutoReset = True
            _timer.Enabled = False

            _dateCreateTotal = Registry.GetValue("HKEY_CURRENT_USER\Software\VB and VBA Program Settings\IsolaEcologica\Sync", "LastTotalCreateDate", Nothing)
            If _dateCreateTotal = Nothing Then _dateCreateTotal = DateAdd(DateInterval.Day, -1, Now)

            AddHandler _timer.Elapsed, AddressOf CreateFilesForSyncDB



            If _propertySetting.ConnectionMode = ConnectionModeEnum.PIPE Then
                'init server Pipe
                InitServerPipe()
            Else
                'init server Shared Folder watcher
                InitWaitSharedFolder()
            End If


            If _nascondi Then
                'importo eventuali file rimasti
                _sync.ImportOldFileCSV()
                'avvio il watcher sulla cartella
                InitWatcher()
                'avvio il timer per la generazione dei file
                _timer.Start()
            End If

        Catch ex As Exception
            My.Application.Log.WriteException(ex)
        End Try



    End Sub

    Dim _dateCreateTotal As Date = New Date(Now.Date.Ticks + My.Settings.TimeEndSync.Ticks)
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="d"></param>
    ''' <remarks></remarks>
    Private Sub CreateFilesForSyncDB(ByVal sender As Object, ByVal d As ElapsedEventArgs)
        Dim tmr = DirectCast(sender, Timer)
        tmr.AutoReset = False
        tmr.Enabled = False

        Try
            'i totali vengono ricreati ad intervalli di tempo  DateInterval.Day
            If New DateTime(d.SignalTime.Ticks) >= _dateCreateTotal Then
                _sync.CreateTotal()
                If _dateCreateTotal < Now Then _dateCreateTotal = Now

                _dateCreateTotal = DateAdd(DateInterval.Day, 1, New Date(_dateCreateTotal.Date.Ticks + My.Settings.TimeEndSync.Ticks))
                Registry.SetValue("HKEY_CURRENT_USER\Software\VB and VBA Program Settings\IsolaEcologica\Sync", "LastTotalCreateDate", _dateCreateTotal)
            End If

            'syncDB
            _sync.PrepareToSync()

            'se collegato via PIPE
            'chiedo ai client di inviarmi i movimenti
            If My.Settings.ConnectionMode = ConnectionModeEnum.PIPE Then
                For Each c As String In _listOfClients
                    SendMessage(c, MessagePipeIDEnum.MessagePipeClientInfo, _
                            1, Pipes.Messagecode.Constants.MESSAGE_CODE_GETMOVIMENTI)
                Next
            Else
                'se vi sono client collegati gli invio i file
                For Each c In ClientsToPing.Values
                    If c.State = SharedFolderListen.SharedFolderStateEnum.Connected Then
                        _sync.SendFileToSync(c.IDPesa)
                    End If
                Next
            End If
        Catch ex As Exception
            My.Application.Log.WriteException(ex)
        Finally
            tmr.AutoReset = True
            tmr.Enabled = True

        End Try


    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Setting_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        NotifyIcon.Dispose()
        _timer.Stop()

        'chiudo server pipe
        EndServerPipe()

        'chiudo shared folder watcher
        EndSharedFolderWatcher()

        My.Application.Log.WriteEntry("Applicazione Fermata")

        Util.Connection.CloseMDB()


    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Setting_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
        If Me.WindowState = FormWindowState.Minimized Then
            _nascondi = True
            Me.Hide()
        End If
    End Sub


    Private Sub Setting_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
        If _nascondi Then Me.Hide()
    End Sub

    Private Sub UscitaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UscitaToolStripMenuItem.Click
        RemoveHandler FileSystemWatcher.Renamed, AddressOf _sync.RenameFileEvent
        Me.Close()
    End Sub

    Private Sub NotifyIcon_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon.MouseClick
        If e.Button = Windows.Forms.MouseButtons.Left Then NotifyIcon.ShowBalloonTip(500)
    End Sub

    Private Sub ImpostazioniToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImpostazioniToolStripMenuItem.Click
        _nascondi = False
        Application.DoEvents()
        Me.Visible = True
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.WindowState = FormWindowState.Normal
        FileSystemWatcher.EnableRaisingEvents = False
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        _propertySetting.Save()
        _sync.RefreshConnection()
        _dateCreateTotal = New Date(Now.Date.Ticks + My.Settings.TimeEndSync.Ticks)
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

    Private Sub InitWatcher()
        Try
            FileSystemWatcher.Path = _propertySetting.FolderNameSync
            FileSystemWatcher.NotifyFilter = IO.NotifyFilters.FileName Or IO.NotifyFilters.CreationTime Or IO.NotifyFilters.Size
            FileSystemWatcher.Filter = _propertySetting.FileNamePattern
            FileSystemWatcher.EnableRaisingEvents = True

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub AvvioToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AvvioToolStripMenuItem.Click
        _dateCreateTotal = Now
    End Sub


#Region "Shared Folder Watcher"
    Private ClientsToPing As Dictionary(Of String, SharedFolderListen) = New Dictionary(Of String, SharedFolderListen)
    ''' <summary>
    ''' avvio il processo background per il monitoraggio delle cartelle condivise con cui il server
    ''' si connette ai clients
    ''' </summary>
    ''' <remarks></remarks>
    Sub InitWaitSharedFolder()
        Dim tc As DataTable = _sync.Clients()

        For Each r As DataRow In tc.Rows
            'creo una classe di tipo 
            Dim folderLister As SharedFolderListen = New SharedFolderListen(My.Settings.SharedFolderName, r.Item("PCName"), r.Item("CodPesa"))
            AddHandler folderLister.ServerConnected, AddressOf ConnectedToPC
            ClientsToPing.Add(r.Item("PCName"), folderLister)
        Next
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

        sb.Append("Sincronizzazione dati IsolaEcologica con le unità mobili." & vbCrLf & _
                                    "Stato Connessioni" & vbCrLf)

        For Each c In ClientsToPing.Values
            sb.Append(String.Format("Unita Mobile {0} - {1}" & vbCrLf, c.PCName, _
                                    If(c.State = SharedFolderListen.SharedFolderStateEnum.Connected, "Connesso", "Non Connesso")))
        Next

        NotifyIcon.BalloonTipText = sb.ToString

    End Sub



    ''' <summary>
    ''' chiusura server watcher
    ''' </summary>
    ''' <remarks></remarks>
    <MethodImplAttribute(MethodImplOptions.Synchronized)> _
    Sub EndSharedFolderWatcher()
        For Each cp In ClientsToPing
            cp.Value.Close()

        Next
        ClientsToPing.Clear()

        My.Application.Log.WriteEntry("Server Shared Folder Stop", TraceEventType.Verbose)
    End Sub
#End Region


#Region "Server Pipe"
    Dim _serverPipe As PipeInstance = Nothing

    Private Sub BackgroundWorker_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BWPipe.DoWork
        _serverPipe.Start("ITAmbienteSync", 3)

        My.Application.Log.WriteEntry("Server Pipe Start", TraceEventType.Verbose)


        'ciclo per mantenere vive le istanze dei server
        Do While Not BWPipe.CancellationPending
            Threading.Thread.Sleep(200)
        Loop

        'chiusura istanze server pipe
        _serverPipe.End()
    End Sub

    ''' <summary>
    ''' Inizializzazione server pipe
    ''' </summary>
    ''' <remarks></remarks>
    Sub InitServerPipe()
        _serverPipe = New PipeInstance

        AddHandler _serverPipe.InstanceConnected, AddressOf InstanceConnected
        AddHandler _serverPipe.InstanceDisconnected, AddressOf InstanceDisconnected
        AddHandler _serverPipe.InstanceNewMessage, AddressOf InstanceNewMessage

        BWPipe.WorkerSupportsCancellation = True
        BWPipe.RunWorkerAsync()



    End Sub



    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="clientID"></param>
    ''' <param name="messageID"></param>
    ''' <param name="messageBody"></param>
    ''' <remarks></remarks>
    Public Sub SendMessage(ByVal clientID As String, ByVal messageID As MessagePipeIDEnum, ByVal messageBody As String, Optional ByVal messageCode As Integer = Nothing)
        Try
            _serverPipe.SendMessage(clientID, New MessagePipe(messageID, messageCode, messageBody))
        Catch ex As Exception
            'todo da gestire l'errore
        End Try

    End Sub

    ''' <summary>
    ''' chiusura server pipe
    ''' </summary>
    ''' <remarks></remarks>
    <MethodImplAttribute(MethodImplOptions.Synchronized)> _
    Sub EndServerPipe()

        If _serverPipe Is Nothing Then Exit Sub

        'chiudo i client
        SendMessage(vbNullString, MessagePipeIDEnum.MessagePipeClientEnd, 1)

        RemoveHandler _serverPipe.InstanceConnected, AddressOf InstanceConnected
        RemoveHandler _serverPipe.InstanceDisconnected, AddressOf InstanceDisconnected
        RemoveHandler _serverPipe.InstanceNewMessage, AddressOf InstanceNewMessage

        BWPipe.CancelAsync()

        Do While BWPipe.IsBusy
            Application.DoEvents()
        Loop


        My.Application.Log.WriteEntry("Server Pipe Stop", TraceEventType.Verbose)

    End Sub


    ''' <summary>
    ''' Ricezione di un nuovo messaggio
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="m"></param>
    ''' <remarks></remarks>
    <MethodImplAttribute(MethodImplOptions.Synchronized)> _
    Private Sub InstanceNewMessage(ByVal sender As Object, ByVal m As MessagePipe)

        If sender.ClientName = "" Then
        End If

        If m.Id = MessagePipeIDEnum.MessagePipeClientName Then

            My.Application.Log.WriteEntry(String.Format("Ricevuto Messaggio dal Client Pipe {0}", sender.ClientName), TraceEventType.Verbose)

            'se vi sono file per il client li invio con un file copy
            Dim t = New Threading.Thread(AddressOf _sync.SendFileToSync)

            t.Start(sender.ClientName)

            'add client name to a list 
            _listOfClients.Add(sender.ClientName)
        ElseIf m.Id = MessagePipeIDEnum.MessagePipeClientInfo Then
            Select Case m.MessageCode
                Case Pipes.Messagecode.Constants.MESSAGE_CODE_FILESYNC
                    Dim s = m.MessageText.Split("|")
                    If s(2) = 1 Then _sync.RemoveFileToSync(s(0), s(1))
            End Select
        End If
    End Sub


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub InstanceConnected(ByVal sender As Object, ByVal e As EventArgs)
        My.Application.Log.WriteEntry("Client Pipe Connesso", TraceEventType.Verbose)
        _listOfClients.Add(sender.ClientName)
    End Sub

    ''' <summary>
    ''' Disconnesione dalla coda PIPE da parte di un Client
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    <MethodImplAttribute(MethodImplOptions.Synchronized)> _
    Private Sub InstanceDisconnected(ByVal sender As Object, ByVal e As EventArgs)
        My.Application.Log.WriteEntry(String.Format("Client Pipe {0} Disconnesso", sender.clientName), TraceEventType.Verbose)
        _listOfClients.Remove(sender.ClientName)
    End Sub

#End Region

End Class
