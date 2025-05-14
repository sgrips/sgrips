Imports System.IO
Imports ITAmbiente.easyrdc.Sync.DataSetMDB

Module SyncClient



    Public gMainForm As ClientSyncUC

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ImportOldFileXML()
        Dim q = From f In System.IO.Directory.GetFiles(My.Settings.FolderNameSync, "*.xml")

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

    Dim objLockImport = New Object
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="fileName"></param>
    ''' <remarks></remarks>
    Public Sub StartImport(ByVal fileName As String)
        SyncLock objLockImport
            Dim fileInfo As IO.FileInfo = My.Computer.FileSystem.GetFileInfo(fileName)
            Dim codPesa As String = vbNullString
            Dim bErr As Boolean = False


            'se la dimensione del file è maggiore di 0 allora inizio il trasferimento
            If fileInfo.Length > 0 And fileInfo.Extension.ToUpper = ".XML" Then

                Threading.Thread.Sleep(200)

                Try
                    My.Application.Log.WriteEntry("ricevuto file" + fileInfo.FullName)
                    Dim t As TotaliPerBadgeDataTable = New TotaliPerBadgeDataTable
                    t.ReadXml(fileInfo.FullName)
                    My.Application.Log.WriteEntry("righe presenti nella tabella" + t.Count.ToString)
                    'importo il file
                Catch ex As Exception
                    My.Application.Log.WriteException(ex)
                    bErr = True
                Finally
                    Try
                        Try
                            If My.Settings.ConnectionMode = PropertySettings.ConnectionModeEnum.SHAREDFOLDER Then
                                'invio un file semaforo per indicare l'avvenuta importazione
                                'ClientID+FlgError+FileName
                                Dim fileCheck As String = String.Format("Movimenti{0:d2}{1}{2}", My.Settings.ClientID, If(bErr, 0, 1), fileInfo.Name)

                                Dim sfileRemote = Path.Combine(String.Format("\\{0}\{1}", _
                                                                              My.Settings.ServerURL, _
                                                                              My.Settings.SharedFolderName), _
                                                                              fileCheck + ".TMP")
                                Dim sw As StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(sfileRemote, True)
                                sw.WriteLine("file sync")
                                sw.Flush()
                                sw.Close()
                                sw.Dispose()
                                Threading.Thread.Sleep(200)
                                My.Computer.FileSystem.RenameFile(sfileRemote, fileCheck + ".result")
                            End If
                        Catch ex As Exception
                            My.Application.Log.WriteException(ex)
                            bErr = True
                        End Try

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

                    Catch ex As Exception
                        My.Application.Log.WriteException(ex)
                    Finally

                        If My.Settings.ConnectionMode = PropertySettings.ConnectionModeEnum.PIPE Then
                            gMainForm.SendMessagePipe(String.Format("{0}|{1}|{2}", My.Settings.ClientID, fileInfo.Name, If(bErr, 0, 1)), _
                                                      Framework.IO.Communication.Pipe.MessagePipeIDEnum.MessagePipeClientInfo, _
                                                       Pipes.Messagecode.Constants.MESSAGE_CODE_GETMOVIMENTI_RESULT)

                        End If
                    End Try
                End Try
            End If
        End SyncLock
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Dim objLockSend As New Object
    Sub SendMovimentiToServer()
        SyncLock objLockSend
            Dim q = From f In System.IO.Directory.GetFiles(My.Settings.FolderNameSync, "*.csv")

            For Each f In q

                Dim fileInfo As IO.FileInfo = My.Computer.FileSystem.GetFileInfo(f)
                Dim codPesa As String = vbNullString
                Dim bErr As Boolean = False


                'se la dimensione del file è maggiore di 0 allora inizio il trasferimento
                If fileInfo.Length > 0 And fileInfo.Extension.ToUpper = ".CSV" Then
                    Try
                        Debug.Print(String.Format("trovato file movimenti {0}", fileInfo.FullName))
                        'spedisco file al server

                        Dim sfileRemote = Path.Combine(String.Format("\\{0}\{1}", _
                                                       My.Settings.ServerURL, _
                                                       My.Settings.SharedFolderName), _
                                                       Path.GetFileNameWithoutExtension(fileInfo.Name) + ".tmp")
                        My.Computer.FileSystem.CopyFile(fileInfo.FullName, sfileRemote, True)
                        Threading.Thread.Sleep(200)
                        My.Computer.FileSystem.RenameFile(sfileRemote, Path.GetFileNameWithoutExtension(fileInfo.Name) + ".csv")
                    Catch ex As Exception
                        My.Application.Log.WriteException(ex)
                        bErr = True
                    Finally
                        Try
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

                        Catch ex As Exception
                            My.Application.Log.WriteException(ex)
                        Finally

                            If My.Settings.ConnectionMode = PropertySettings.ConnectionModeEnum.PIPE Then
                                gMainForm.SendMessagePipe(String.Format("{0}|{1}|{2}", My.Settings.ClientID, fileInfo.Name, If(bErr, 0, 1)), _
                                                          Framework.IO.Communication.Pipe.MessagePipeIDEnum.MessagePipeClientInfo, _
                                                           Pipes.Messagecode.Constants.MESSAGE_CODE_GETMOVIMENTI_RESULT)

                            End If
                        End Try

                    End Try
                End If
            Next

        End SyncLock

    End Sub

End Module
