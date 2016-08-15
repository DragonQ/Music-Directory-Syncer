#Region " Namespaces "
Imports MusicFolderSyncer.Codec.CodecType
Imports MusicFolderSyncer.Toolkit
Imports MusicFolderSyncer.Logger.DebugLogLevel
Imports System.Windows.Forms
Imports System.Environment
Imports System.IO
Imports System.Threading
#End Region

Public Class TrayApp
    Inherits ApplicationContext

#Region " Declarations "
    Private WithEvents Tray As NotifyIcon
    Private WithEvents MainMenu As ContextMenuStrip
    Private WithEvents mnuViewLogFile, mnuNewSync, mnuStatus, mnuEditSyncSettings, mnuExit As ToolStripMenuItem
    Private WithEvents mnuSep1, mnuSep2, mnuSep3 As ToolStripSeparator
    Private Const BalloonTime As Int32 = 8

    Private WithEvents FileWatcher As FileSystemWatcher
    Private FileID As Int32 = 0
#End Region


    ' TO DO:
    ' - Add button in context menu for enable/disable sync!

#Region " New "
    Public Sub New(LaunchNewSyncWindow As Boolean)

        'Initialize the menus
        mnuStatus = New ToolStripMenuItem("Syncer is not active")
        mnuSep1 = New ToolStripSeparator()
        mnuEditSyncSettings = New ToolStripMenuItem("Edit sync settings")
        mnuViewLogFile = New ToolStripMenuItem("View log file")
        mnuSep2 = New ToolStripSeparator()
        mnuNewSync = New ToolStripMenuItem("Create new sync")
        mnuSep3 = New ToolStripSeparator()
        mnuExit = New ToolStripMenuItem("Exit")
        mnuStatus.Enabled = False
        mnuEditSyncSettings.Enabled = False
        MainMenu = New ContextMenuStrip
        MainMenu.Items.AddRange(New ToolStripItem() {mnuStatus, mnuSep1, mnuEditSyncSettings, mnuViewLogFile, mnuSep2, mnuNewSync, mnuSep3, mnuExit})

        'Initialize the notification area icon
        Tray = New NotifyIcon
        Tray.Icon = My.Resources.Tray_Icon
        Tray.ContextMenuStrip = MainMenu
        Tray.Text = ApplicationName

        'Either display the Create New Sync window or start the file system watcher for background syncing
        If LaunchNewSyncWindow Then
            ShowNewSyncWindow()
        Else
            StartWatcher()
            mnuEditSyncSettings.Enabled = True
            Tray.Visible = True
        End If

    End Sub
#End Region

#Region " Event Handlers "
    Private Sub AppContext_ThreadExit(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ThreadExit
        Tray.Visible = False 'Guarantees that the icon will not linger.
    End Sub

    Private Sub mnuViewLogFile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuViewLogFile.Click
        Process.Start(MyLogFilePath)
    End Sub

    Private Sub mnuEditSyncSettings_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuEditSyncSettings.Click
        ShowEditSyncSettingsWindow()
    End Sub

    Private Sub mnuExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        ExitApplication()
    End Sub

    Public Sub ExitApplication()
        'Perform any clean-up here then exit the application
        MyLog.Write("===============================================================")
        MyLog.Write("  PROGRAM CLEAN EXIT")
        MyLog.Write("===============================================================")
        MyLog.Write("")
        Tray.Visible = False
        ExitThread() 'IF THIS EVER CAUSES ISSUES, USE THIS INSTEAD: Forms.Application.Exit()
    End Sub

    Private Sub mnuNewSync_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuNewSync.Click
        ShowNewSyncWindow()
    End Sub

#End Region

#Region " Open Windows "
    Private Sub ShowEditSyncSettingsWindow()

        Dim MyEditSyncSettingsWindow As New EditSyncSettingsWindow
        Tray.Visible = False
        MyEditSyncSettingsWindow.ShowDialog()

        Tray.Visible = True

        If MyEditSyncSettingsWindow.DialogResult = True Then
            Tray.ShowBalloonTip(BalloonTime, ApplicationName, "Sync settings updated.", ToolTipIcon.Info)
        End If

    End Sub

    Private Sub ShowNewSyncWindow()

        Dim MyNewSyncWindow As New NewSyncWindow
        Tray.Visible = False
        MyNewSyncWindow.ShowDialog()

        Tray.Visible = True

        If MyNewSyncWindow.DialogResult = True Then
            ' Sync was successfully set up
            mnuEditSyncSettings.Enabled = True
            If MyGlobalSyncSettings.SyncIsEnabled Then
                Dim WatcherStartResult As ReturnObject = StartWatcher()

                If WatcherStartResult.Success Then
                    mnuStatus.Text = "Syncer is active"
                    Tray.ShowBalloonTip(BalloonTime, ApplicationName, "Syncer active.", ToolTipIcon.Info)
                Else
                    System.Windows.MessageBox.Show("Error starting background syncer!" & NewLine & NewLine & WatcherStartResult.ErrorMessage, "Background Syncer Error!", MessageBoxButton.OK, MessageBoxImage.Error)
                End If
            Else
                mnuStatus.Text = "Syncer is not active"
                Tray.ShowBalloonTip(BalloonTime, ApplicationName, "Syncer disabled.", ToolTipIcon.Info)
            End If
        Else ' User closed the window before sync was completed
            mnuStatus.Text = "Syncer is not active"
            Tray.ShowBalloonTip(BalloonTime, ApplicationName, "Syncer not set up.", ToolTipIcon.Info)
        End If

    End Sub
#End Region

#Region " File System Watcher "
    Private Function StartWatcher() As ReturnObject

        If Not Directory.Exists(MyGlobalSyncSettings.SourceDirectory) Then
            Return New ReturnObject(False, "Directory """ + MyGlobalSyncSettings.SourceDirectory + """ does not exist!", Nothing)
        End If

        Try
            FileWatcher = New FileSystemWatcher(MyGlobalSyncSettings.SourceDirectory)
            FileWatcher.IncludeSubdirectories = True
            FileWatcher.NotifyFilter = NotifyFilters.FileName Or NotifyFilters.LastWrite Or NotifyFilters.DirectoryName
            FileWatcher.EnableRaisingEvents = True

            AddHandler FileWatcher.Changed, AddressOf FileChanged
            AddHandler FileWatcher.Created, AddressOf FileChanged
            AddHandler FileWatcher.Renamed, AddressOf FileRenamed
            AddHandler FileWatcher.Deleted, AddressOf FileChanged

            MyLog.Write("File system watcher started (monitoring directory """ & MyGlobalSyncSettings.SourceDirectory & """ for audio files)", Information)
            mnuStatus.Text = "Syncer is active"
            If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, ApplicationName, "Syncer active.", ToolTipIcon.Info)
        Catch ex As Exception
            Return New ReturnObject(False, ex.Message, Nothing)
        End Try

        Return New ReturnObject(True, "", Nothing)

    End Function

    Private Sub StopWatcher()

        FileWatcher.Dispose()
        MyGlobalSyncSettings.SyncIsEnabled = False
        mnuStatus.Text = "Syncer is not active"

        Dim MyResult As ReturnObject = SaveSyncSettings()

        If MyResult.Success Then
            MyLog.Write("Syncer stopped.", Warning)
            If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, ApplicationName, "Syncer has been disabled.", ToolTipIcon.Info)
        Else
            MyLog.Write("Could not update sync settings. Error: " & MyResult.ErrorMessage, Warning)
            If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, ApplicationName, "Syncer has been disabled but settings file could not be updated.", ToolTipIcon.Error)
        End If

    End Sub

    Private Sub FileRenamed(ByVal sender As Object, ByVal e As RenamedEventArgs)

        If FilterMatch(e.Name) Then
            MyLog.Write("File renamed: " & e.FullPath, Information)
            Dim MyFileParser As New FileParser(MyGlobalSyncSettings, FileID, e.FullPath)
            RenameInSyncFolder(MyFileParser, e.OldFullPath)
            MyFileParser = Nothing
        End If

        If Interlocked.Equals(FileID, MaxFileID) Then
            Interlocked.Add(FileID, -MaxFileID)
        Else
            Interlocked.Increment(FileID)
        End If

    End Sub

    Private Sub FileChanged(ByVal sender As Object, ByVal e As System.IO.FileSystemEventArgs)
        'Handles changed and new files

        If FilterMatch(e.Name) Then
            Dim MyFileParser As New FileParser(MyGlobalSyncSettings, FileID, e.FullPath)
            Select Case e.ChangeType
                Case Is = IO.WatcherChangeTypes.Changed
                    MyLog.Write("File changed: " & e.FullPath, Information)
                    DeleteInSyncFolder(MyFileParser, e.FullPath)
                    Dim Result As ReturnObject = MyFileParser.TransferToSyncFolder(MySyncSettings.GetWatcherCodecs)

                    If Result.Success Then
                        If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, "File processed:", e.FullPath.Substring(MySyncSettings.SourceDirectory.Length), ToolTipIcon.Info)
                    Else
                        If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, "File processing failed:", e.FullPath.Substring(MySyncSettings.SourceDirectory.Length), ToolTipIcon.Error)
                    End If
                Case Is = IO.WatcherChangeTypes.Created
                    MyLog.Write("File created: " & e.FullPath, Information)
                    Dim Result As ReturnObject = MyFileParser.TransferToSyncFolder(MySyncSettings.GetWatcherCodecs)

                    If Result.Success Then
                        If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, "File processed:", e.FullPath.Substring(MySyncSettings.SourceDirectory.Length), ToolTipIcon.Info)
                    Else
                        If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, "File processing failed:", e.FullPath.Substring(MySyncSettings.SourceDirectory.Length), ToolTipIcon.Error)
                    End If
                Case Is = IO.WatcherChangeTypes.Deleted
                    MyLog.Write("File deleted: " & e.FullPath, Information)
                    DeleteInSyncFolder(MyFileParser, e.FullPath)
            End Select
            MyFileParser = Nothing
        End If

        If Interlocked.Equals(FileID, MaxFileID) Then
            Interlocked.Add(FileID, -MaxFileID)
        Else
            Interlocked.Increment(FileID)
        End If

    End Sub

    Private Shared Function FilterMatch(FileName As String) As Boolean

        Dim Match As Boolean = False
        Dim FileExtension As String = Path.GetExtension(FileName).ToLower(EnglishGB)

        For Each Filter As String In MySyncSettings.GetFileExtensions()
            If FileExtension = Filter.ToLower(EnglishGB) Then
                Return True
            End If
        Next

        Return False

    End Function

    Private Sub DeleteInSyncFolder(ByRef MyFileParser As FileParser, ByVal FilePath As String)

        Try
            Dim FileCodec As Codec = MyFileParser.CheckFileCodec(MySyncSettings.GetWatcherCodecs)

            If Not FileCodec Is Nothing Then
                'File was meant to be synced, which means we now need to delete the synced version

                Dim SyncFilePath As String = MySyncSettings.SyncDirectory & FilePath.Substring(MyGlobalSyncSettings.SourceDirectory.Length)

                If MySyncSettings.TranscodeLosslessFiles AndAlso FileCodec.CompressionType = Lossless Then 'Need to replace extension with .ogg
                    Dim TranscodedFilePath As String = Path.Combine(Path.GetDirectoryName(SyncFilePath), Path.GetFileNameWithoutExtension(SyncFilePath)) &
                                                    MySyncSettings.Encoder.GetFileExtensions(0)
                    SyncFilePath = TranscodedFilePath
                End If

                'Delete file if it exists in sync folder
                If File.Exists(SyncFilePath) Then
                    File.Delete(SyncFilePath)
                    MyLog.Write("...file in sync folder deleted: """ & SyncFilePath.Substring(MySyncSettings.SyncDirectory.Length) & """.", Information)
                    If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, "File Deleted:", SyncFilePath.Substring(MySyncSettings.SyncDirectory.Length), ToolTipIcon.Info)
                Else
                    MyLog.Write("...file doesn't exist in sync folder: """ & SyncFilePath.Substring(MySyncSettings.SyncDirectory.Length) & """.", Information)
                End If
            Else
                Throw New Exception("File was being watched but could not determine its codec.")
            End If
        Catch ex As Exception
            MyLog.Write("...couldn't delete file in sync folder. Exception: " & ex.Message, Warning)
        End Try

    End Sub

    Private Sub RenameInSyncFolder(ByRef MyFileParser As FileParser, ByVal OldFilePath As String)

        Try
            Dim FileCodec As Codec = MyFileParser.CheckFileCodec(MySyncSettings.GetWatcherCodecs)
            Dim NewFilePath As String = MyFileParser.FilePath

            If Not FileCodec Is Nothing Then

                If MyFileParser.CheckFileForSync(FileCodec) Then
                    Dim SyncFilePath As String = MySyncSettings.SyncDirectory & NewFilePath.Substring(MyGlobalSyncSettings.SourceDirectory.Length)
                    Dim OldSyncFilePath As String = MySyncSettings.SyncDirectory & OldFilePath.Substring(MyGlobalSyncSettings.SourceDirectory.Length)

                    If MySyncSettings.TranscodeLosslessFiles AndAlso FileCodec.CompressionType = Lossless Then 'Need to replace extension with .ogg
                        Dim TempString As String = Path.Combine(Path.GetDirectoryName(SyncFilePath), Path.GetFileNameWithoutExtension(SyncFilePath)) &
                            MySyncSettings.Encoder.GetFileExtensions(0)
                        SyncFilePath = TempString
                        TempString = Path.Combine(Path.GetDirectoryName(OldSyncFilePath), Path.GetFileNameWithoutExtension(OldSyncFilePath)) &
                            MySyncSettings.Encoder.GetFileExtensions(0)
                        OldSyncFilePath = TempString
                    End If

                    If File.Exists(OldSyncFilePath) Then
                        File.Move(OldSyncFilePath, SyncFilePath)
                    Else
                        MyLog.Write("...old file doesn't exist in sync folder: """ & OldSyncFilePath & """, creating now...", Warning)

                        If MySyncSettings.TranscodeLosslessFiles AndAlso FileCodec.CompressionType = Lossless Then 'Need to transcode file
                            MyLog.Write("...transcoding file to " & MySyncSettings.Encoder.Name & "...", Debug)
                            MyFileParser.TranscodeFile(SyncFilePath)
                        Else
                            Directory.CreateDirectory(Path.GetDirectoryName(SyncFilePath))
                            File.Copy(NewFilePath, SyncFilePath, True)
                        End If

                        MyLog.Write("...successfully added file to sync folder.", Information)
                        If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, "File Processed:", NewFilePath.Substring(MyGlobalSyncSettings.SourceDirectory.Length),
                                                ToolTipIcon.Info)
                    End If

                    MyLog.Write("...successfully renamed file in sync folder.", Information)
                    If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, "File Renamed:", NewFilePath.Substring(MyGlobalSyncSettings.SourceDirectory.Length),
                                                ToolTipIcon.Info)
                End If

            End If
        Catch ex As Exception
            MyLog.Write("...failed to add file to sync folder. Exception: " & ex.Message, Warning)
        End Try

    End Sub
#End Region

End Class