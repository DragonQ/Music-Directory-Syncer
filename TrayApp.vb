#Region " Namespaces "
Imports MusicFolderSyncer.Toolkit
Imports MusicFolderSyncer.Logger.LogLevel
Imports System.Windows.Forms
Imports System.Environment
Imports System.IO
Imports System.Threading
#End Region

Public Class TrayApp
    Inherits ApplicationContext

#Region " Declarations "
    Private ApplicationExitPending As Boolean = False

    Private WithEvents Tray As NotifyIcon
    Private WithEvents MainMenu As ContextMenuStrip
    Private WithEvents mnuEditSyncSettings, mnuEnableSync, mnuExit, mnuNewSync, mnuReapplySyncs, mnuStatus, mnuViewLogFile As ToolStripMenuItem
    Private WithEvents mnuSep1, mnuSep2, mnuSep3 As ToolStripSeparator
    Private Const BalloonTime As Int32 = 8 * 1000

    Private WithEvents DirectoryWatcher As FileSystemWatcher
    Private WithEvents FileWatcher As FileSystemWatcher
    Private FileID As Int32 = 0

    Private MySyncer As SyncerInitialiser = Nothing
    Private SyncTimer As New Stopwatch()
    Private SyncInProgress As Boolean = False
#End Region

#Region " New "
    Public Sub New(LaunchNewSyncWindow As Boolean)

        'Initialize the menus
        mnuStatus = New ToolStripMenuItem("Syncer is not active")
        mnuSep1 = New ToolStripSeparator()
        mnuEnableSync = New ToolStripMenuItem("Enable sync")
        mnuEditSyncSettings = New ToolStripMenuItem("Edit sync settings")
        mnuViewLogFile = New ToolStripMenuItem("View log file")
        mnuSep2 = New ToolStripSeparator()
        mnuReapplySyncs = New ToolStripMenuItem("Re-apply all syncs")
        mnuNewSync = New ToolStripMenuItem("Create new sync")
        mnuSep3 = New ToolStripSeparator()
        mnuExit = New ToolStripMenuItem("Exit")
        mnuStatus.Enabled = False
        EnableDisableControls(False)
        MainMenu = New ContextMenuStrip
        MainMenu.Items.AddRange(New ToolStripItem() {mnuStatus,
                                                     mnuSep1, mnuEnableSync, mnuEditSyncSettings, mnuViewLogFile,
                                                     mnuSep2, mnuReapplySyncs, mnuNewSync,
                                                     mnuSep3, mnuExit})

        'Initialize the notification area icon
        Tray = New NotifyIcon
        Tray.Icon = My.Resources.Tray_Icon
        Tray.ContextMenuStrip = MainMenu
        Tray.Text = ApplicationName

        'Either display the Create New Sync window or start the file system watcher for background syncing
        If LaunchNewSyncWindow Then
            ShowNewSyncWindow()
        Else
            If UserGlobalSyncSettings.SyncIsEnabled Then
                Dim WatcherStartResult As ReturnObject = StartWatcher()
                If Not WatcherStartResult.Success Then
                    Dim ErrorMessage As String = "File system watcher could not be started. " & NewLine & NewLine & WatcherStartResult.ErrorMessage
                    MyLog.Write(ErrorMessage, Fatal)
                    System.Windows.MessageBox.Show(ErrorMessage, "File System Watcher Error", MessageBoxButton.OK, MessageBoxImage.Error)
                    ExitApplication()
                End If
            Else
                Dim WatcherStopResult As ReturnObject = StopWatcher()
                If Not WatcherStopResult.Success Then
                    Dim ErrorMessage As String = "File system watcher could not be stopped. " & NewLine & NewLine & WatcherStopResult.ErrorMessage
                    MyLog.Write(ErrorMessage, Warning)
                    System.Windows.MessageBox.Show(ErrorMessage, "File System Watcher Error", MessageBoxButton.OK, MessageBoxImage.Error)
                End If
            End If
            EnableDisableControls(True)
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

    Private Sub mnuEnableSync_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuEnableSync.Click

        Dim Success As Boolean = True

        'Toggle user setting for sync being enabled
        If UserGlobalSyncSettings.SyncIsEnabled Then 'Disable sync
            Dim WatcherStopResult As ReturnObject = StopWatcher()
            If Not WatcherStopResult.Success Then
                Dim ErrorMessage As String = "File system watcher could not be stopped. " & NewLine & NewLine & WatcherStopResult.ErrorMessage
                MyLog.Write(ErrorMessage, Warning)
                If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, ApplicationName, ErrorMessage, ToolTipIcon.Error)
                Success = False
            End If
        Else 'Enable sync
            Dim WatcherStartResult As ReturnObject = StartWatcher()
            If Not WatcherStartResult.Success Then
                Dim ErrorMessage As String = "File system watcher could not be started. " & NewLine & NewLine & WatcherStartResult.ErrorMessage
                MyLog.Write(ErrorMessage, Warning)
                If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, ApplicationName, ErrorMessage, ToolTipIcon.Error)
                Success = False
            End If
        End If

        If Success Then 'Toggle sync enabled setting and save settings
            UserGlobalSyncSettings.SyncIsEnabled = Not UserGlobalSyncSettings.SyncIsEnabled

            Dim SaveResult As ReturnObject = SaveSyncSettings(UserGlobalSyncSettings)
            If Not SaveResult.Success Then
                Dim ErrorMessage As String = "New sync settings could not be saved. " & NewLine & NewLine & SaveResult.ErrorMessage
                MyLog.Write(ErrorMessage, Warning)
                If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, "Error", "Could not save new sync settings: " & NewLine & SaveResult.ErrorMessage, ToolTipIcon.Error)
            End If
        End If

    End Sub

    Private Sub ApplyEnableSync()

        If UserGlobalSyncSettings.SyncIsEnabled Then 'Enable sync
            Dim WatcherStartResult As ReturnObject = StartWatcher()
            If Not WatcherStartResult.Success Then
                Dim ErrorMessage As String = "File system watcher could not be started. " & NewLine & NewLine & WatcherStartResult.ErrorMessage
                MyLog.Write(ErrorMessage, Warning)
                If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, ApplicationName, ErrorMessage, ToolTipIcon.Error)
            End If
        Else 'Disable sync
            Dim WatcherStopResult As ReturnObject = StopWatcher()
            If Not WatcherStopResult.Success Then
                Dim ErrorMessage As String = "File system watcher could not be stopped. " & NewLine & NewLine & WatcherStopResult.ErrorMessage
                MyLog.Write(ErrorMessage, Warning)
                If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, ApplicationName, ErrorMessage, ToolTipIcon.Error)
            End If
        End If

    End Sub

    Private Sub mnuReapplySyncs_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuReapplySyncs.Click

        'Double check that the user wants to do this
        If System.Windows.MessageBox.Show("Are you sure you want to manually re-create all of your sync directories? This can solve file mis-match issues but will take a long time.", "Re-Apply All Syncs",
                                   MessageBoxButton.OKCancel, MessageBoxImage.Error) = MessageBoxResult.Cancel Then
            Exit Sub
        End If

        EnableDisableControls(False)

        'Stop the watcher if it's running
        Dim WatcherStopResult As ReturnObject = StopWatcher(False)
        If Not WatcherStopResult.Success Then
            Dim ErrorMessage As String = "File system watcher could not be stopped. " & NewLine & NewLine & WatcherStopResult.ErrorMessage
            MyLog.Write(ErrorMessage, Warning)
            System.Windows.MessageBox.Show(ErrorMessage, "File System Watcher Error", MessageBoxButton.OK, MessageBoxImage.Error)
            EnableDisableControls(True)
            Exit Sub
        End If

        SyncInProgress = True
        MyLog.Write("Starting re-sync of all sync directories", Information)

        'Create syncer initialiser
        MySyncer = New SyncerInitialiser(UserGlobalSyncSettings, 500)

        'Set callback functions for the SyncBackgroundWorker
        MySyncer.AddProgressCallback(AddressOf SyncFolderProgressChanged)
        MySyncer.AddCompletionCallback(AddressOf SyncFolderCompleted)

        'Start sync
        If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, "Processing files...", "Right-click or hover over the application icon to see progress.", ToolTipIcon.Info)
        mnuStatus.Text = "File processing initialising..."
        SyncTimer.Start()
        MySyncer.InitialiseSync()

    End Sub

    Private Sub mnuEditSyncSettings_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuEditSyncSettings.Click
        ShowEditSyncSettingsWindow()
    End Sub

    Private Sub mnuExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuExit.Click

        If SyncInProgress Then
            If System.Windows.MessageBox.Show("Are you sure you want to exit? Your sync directories will be incomplete!", "Sync in progress!",
                                   MessageBoxButton.OKCancel, MessageBoxImage.Error) = MessageBoxResult.OK Then
                ApplicationExitPending = True
                CancelSync()
            End If
        Else
            ExitApplication()
        End If

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

        Dim MyEditSyncSettingsWindow As New EditSyncSettingsWindow(UserGlobalSyncSettings)
        Tray.Visible = False
        MyEditSyncSettingsWindow.ShowDialog()

        Tray.Visible = True

        If MyEditSyncSettingsWindow.DialogResult = True Then
            Tray.ShowBalloonTip(BalloonTime, ApplicationName, "Sync settings updated.", ToolTipIcon.Info)
        End If

    End Sub

    Private Sub ShowNewSyncWindow()

        Dim MyGlobalSyncSettings As GlobalSyncSettings
        Dim NewSync As Boolean = True
        If UserGlobalSyncSettings Is Nothing Then
            MyGlobalSyncSettings = DefaultGlobalSyncSettings
        Else
            MyGlobalSyncSettings = UserGlobalSyncSettings
            NewSync = False
        End If
        Dim MyNewSyncWindow As New NewSyncWindow(MyGlobalSyncSettings, NewSync)
        Tray.Visible = False
        MyNewSyncWindow.ShowDialog()

        Tray.Visible = True

        If MyNewSyncWindow.DialogResult = True Then
            'Sync was successfully set up
            ApplyEnableSync()
            EnableDisableControls(True)
        Else 'User closed the window before sync was completed
            mnuStatus.Text = "Syncer is not active"
            Tray.ShowBalloonTip(BalloonTime, ApplicationName, "Syncer not set up.", ToolTipIcon.Info)
            mnuNewSync.Enabled = True
        End If

    End Sub

    Private Sub EnableDisableControls(Enable As Boolean)

        'Group box controls
        mnuEditSyncSettings.Enabled = Enable
        mnuEnableSync.Enabled = Enable
        mnuReapplySyncs.Enabled = Enable
        mnuNewSync.Enabled = Enable

    End Sub
#End Region

#Region " File System Watcher "
    Private Function StartWatcher() As ReturnObject

        If Not Directory.Exists(UserGlobalSyncSettings.SourceDirectory) Then
            Return New ReturnObject(False, "Directory """ + UserGlobalSyncSettings.SourceDirectory + """ does not exist!", Nothing)
        End If

        Try
            'Create file watcher
            FileWatcher = New FileSystemWatcher(UserGlobalSyncSettings.SourceDirectory)
            FileWatcher.IncludeSubdirectories = True
            FileWatcher.NotifyFilter = NotifyFilters.FileName Or NotifyFilters.LastWrite
            FileWatcher.EnableRaisingEvents = True

            'Add handlers for file watcher
            AddHandler FileWatcher.Changed, AddressOf FileChanged
            AddHandler FileWatcher.Created, AddressOf FileChanged
            AddHandler FileWatcher.Renamed, AddressOf FileRenamed
            AddHandler FileWatcher.Deleted, AddressOf FileChanged

            'Create directory watcher
            DirectoryWatcher = New FileSystemWatcher(UserGlobalSyncSettings.SourceDirectory)
            DirectoryWatcher.IncludeSubdirectories = True
            DirectoryWatcher.NotifyFilter = NotifyFilters.DirectoryName
            DirectoryWatcher.EnableRaisingEvents = True

            'Add handlers for directory watcher
            AddHandler DirectoryWatcher.Changed, AddressOf FileChanged
            AddHandler DirectoryWatcher.Created, AddressOf FileChanged
            AddHandler DirectoryWatcher.Renamed, AddressOf DirectoryRenamed
            AddHandler DirectoryWatcher.Deleted, AddressOf FileChanged

            mnuStatus.Text = "Syncer is active"
            mnuEnableSync.Text = "Disable sync"
            MyLog.Write("File system watcher started (monitoring directory """ & UserGlobalSyncSettings.SourceDirectory & """ for audio files)", Information)
            If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, ApplicationName, "Syncer active.", ToolTipIcon.Info)

            Return New ReturnObject(True, Nothing)
        Catch ex As Exception
            Return New ReturnObject(False, ex.Message, Nothing)
        End Try

    End Function

    Private Function StopWatcher(Optional ShowTooltip As Boolean = True) As ReturnObject

        If FileWatcher IsNot Nothing Then FileWatcher.Dispose()

        mnuStatus.Text = "Syncer is not active"
        mnuEnableSync.Text = "Enable sync"
        MyLog.Write("File system watcher stopped.", Information)
        If ShowTooltip AndAlso Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, ApplicationName, "Syncer disabled.", ToolTipIcon.Info)

        Return New ReturnObject(True, Nothing)

    End Function

    Private Sub DirectoryRenamed(ByVal sender As Object, ByVal e As RenamedEventArgs)

        MyLog.Write(FileID, "Directory renamed: " & e.FullPath, Information)

        Dim NewDirectory As New DirectoryInfo(e.FullPath)
        For Each SubDirectory As DirectoryInfo In NewDirectory.GetDirectories("*", SearchOption.TopDirectoryOnly)
            Dim NewDirName = Path.Combine(e.FullPath.Replace(UserGlobalSyncSettings.SourceDirectory & "\", ""), SubDirectory.Name)
            Dim OldDirName = Path.Combine(e.OldFullPath.Replace(UserGlobalSyncSettings.SourceDirectory & "\", ""), SubDirectory.Name)
            DirectoryRenamed(sender, New RenamedEventArgs(WatcherChangeTypes.Renamed, UserGlobalSyncSettings.SourceDirectory, NewDirName, OldDirName))
        Next
        For Each NewFile As FileInfo In NewDirectory.GetFiles("*", SearchOption.TopDirectoryOnly)
            Dim NewFileName = Path.Combine(e.FullPath.Replace(UserGlobalSyncSettings.SourceDirectory & "\", ""), NewFile.Name)
            Dim OldFileName = Path.Combine(e.OldFullPath.Replace(UserGlobalSyncSettings.SourceDirectory & "\", ""), NewFile.Name)
            FileRenamed(sender, New RenamedEventArgs(WatcherChangeTypes.Renamed, UserGlobalSyncSettings.SourceDirectory, NewFileName, OldFileName))
        Next

    End Sub

    Private Sub FileRenamed(ByVal sender As Object, ByVal e As RenamedEventArgs)

        If FilterMatch(e.Name) Then
            'This is a file that we are watching
            MyLog.Write(FileID, "File renamed: " & e.FullPath, Information)
            Using MyFileParser As New FileParser(UserGlobalSyncSettings, FileID, e.FullPath)
                Dim Result As ReturnObject = MyFileParser.RenameInSyncFolder(e.OldFullPath)
                If Result.Success Then
                    If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, "File renamed:", e.FullPath.Substring(UserGlobalSyncSettings.SourceDirectory.Length), ToolTipIcon.Info)
                Else
                    If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, "File rename failed:", e.FullPath.Substring(UserGlobalSyncSettings.SourceDirectory.Length), ToolTipIcon.Error)
                End If
            End Using
        End If

        If Interlocked.Equals(FileID, MaxFileID) Then
            Interlocked.Add(FileID, -MaxFileID)
        Else
            Interlocked.Increment(FileID)
        End If

    End Sub

    Private Sub FileChanged(ByVal sender As Object, ByVal e As System.IO.FileSystemEventArgs)
        'Handles changed and new files

        'Only process this file if it is truly a file and not a directory
        If sender Is FileWatcher Then
            If (Not Directory.Exists(e.FullPath)) AndAlso (FilterMatch(e.Name)) Then
                Using MyFileParser As New FileParser(UserGlobalSyncSettings, FileID, e.FullPath)
                    Select Case e.ChangeType
                        Case Is = IO.WatcherChangeTypes.Changed
                            'File was changed, so re-transfer to all sync directories
                            MyLog.Write(FileID, "Source file changed: " & e.FullPath, Information)
                            Dim Result As ReturnObject = MyFileParser.DeleteInSyncFolder()

                            If Result.Success Then
                                Result = MyFileParser.TransferToSyncFolder()
                            End If

                            If Result.Success Then
                                If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, "File processed:", e.FullPath.Substring(UserGlobalSyncSettings.SourceDirectory.Length), ToolTipIcon.Info)
                            Else
                                If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, "File processing failed:", e.FullPath.Substring(UserGlobalSyncSettings.SourceDirectory.Length), ToolTipIcon.Error)
                            End If
                        Case Is = IO.WatcherChangeTypes.Created
                            'File was created, so transfer to all sync directories
                            MyLog.Write(FileID, "Source file created: " & e.FullPath, Information)
                            Dim Result As ReturnObject = MyFileParser.TransferToSyncFolder()

                            If Result.Success Then
                                If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, "File processed:", e.FullPath.Substring(UserGlobalSyncSettings.SourceDirectory.Length), ToolTipIcon.Info)
                            Else
                                If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, "File processing failed:", e.FullPath.Substring(UserGlobalSyncSettings.SourceDirectory.Length), ToolTipIcon.Error)
                            End If
                        Case Is = IO.WatcherChangeTypes.Deleted
                            'File was deleted, so delete all matching synced files
                            MyLog.Write(FileID, "Source file deleted: " & e.FullPath, Information)
                            Dim Result As ReturnObject = MyFileParser.DeleteInSyncFolder()

                            If Result.Success Then
                                If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, "File deleted:", e.FullPath.Substring(UserGlobalSyncSettings.SourceDirectory.Length), ToolTipIcon.Info)
                            Else
                                If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, "File deletion failed:", e.FullPath.Substring(UserGlobalSyncSettings.SourceDirectory.Length), ToolTipIcon.Error)
                            End If
                    End Select
                End Using
            End If
        ElseIf sender Is DirectoryWatcher Then
            If e.ChangeType = IO.WatcherChangeTypes.Deleted Then
                'Directory was deleted, so delete all matching synced directories
                Using MyFileParser As New FileParser(UserGlobalSyncSettings, FileID, e.FullPath)
                    MyLog.Write(FileID, "Source directory deleted: " & e.FullPath, Information)
                    Dim Result As ReturnObject = MyFileParser.DeleteDirectoryInSyncFolder()

                    If Result.Success Then
                        If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, "Directory deleted:", e.FullPath.Substring(UserGlobalSyncSettings.SourceDirectory.Length), ToolTipIcon.Info)
                    Else
                        If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, "Directory deletion failed:", e.FullPath.Substring(UserGlobalSyncSettings.SourceDirectory.Length), ToolTipIcon.Error)
                    End If
                End Using
            End If
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
        Dim SyncSettings As SyncSettings() = UserGlobalSyncSettings.GetSyncSettings()

        For Each SyncSetting As SyncSettings In SyncSettings
            For Each Filter As String In SyncSetting.GetFileExtensions()
                If FileExtension = Filter.ToLower(EnglishGB) Then
                    Return True
                End If
            Next
        Next

        Return False

    End Function
#End Region

#Region " Re-Sync Callbacks "

    Private Sub CancelSync()
        If Not MySyncer Is Nothing Then
            MySyncer.SyncBackgroundWorker.CancelAsync()
        End If
    End Sub

    Private Sub SyncFolderProgressChanged(sender As Object, e As ProgressChangedEventArgs)

        mnuStatus.Text = "Processing files (" & e.ProgressPercentage & "% complete)"
        Tray.Text = "Processing files (" & e.ProgressPercentage & "% complete)"

    End Sub

    Private Sub SyncFolderCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs)

        SyncTimer.Stop()
        MySyncer = Nothing
        Tray.Text = "Processing files (100% complete)"

        'If the sync task was cancelled, then we need to force-close all instances of ffmpeg and exit application
        If e.Cancelled Then
            Try
                MyLog.Write("Sync cancelled, now force-closing ffmpeg instances...", Warning)

                Dim taskkill As New ProcessStartInfo("taskkill")
                taskkill.CreateNoWindow = True
                taskkill.UseShellExecute = False
                taskkill.Arguments = " /F /IM " & Path.GetFileName(UserGlobalSyncSettings.ffmpegPath) & " /T"
                Dim taskkillProcess As Process = Process.Start(taskkill)
                taskkillProcess.WaitForExit()

                MyLog.Write("...done!", Warning)
            Catch ex As Exception
                MyLog.Write("...failed! There may be lingering ffmpeg instances.", Warning)
            End Try

            SyncInProgress = False
            If ApplicationExitPending Then ExitApplication()
        Else 'Task completed successfully
            SyncInProgress = False

            If Not e.Result Is Nothing Then
                Dim Result As ReturnObject = CType(e.Result, ReturnObject)

                If Result.Success Then
                    'Work out size of sync folder
                    Dim SyncSize As Double = CType(Result.MyObject, Double) / (2 ^ 20) ' Convert to MiB
                    Dim SyncSizeString As String = ""
                    If SyncSize > 1024 Then ' Directory size is greater than 1 GiB
                        SyncSizeString = String.Format(EnglishGB, "{0:0.0}", SyncSize / (2 ^ 10)) & " GiB"
                    Else
                        SyncSizeString = String.Format(EnglishGB, "{0:0.0}", SyncSize) & " MiB"
                    End If

                    'Work out how long the sync took
                    Dim SecondsTaken As Int64 = CInt(Math.Round(SyncTimer.ElapsedMilliseconds / 1000, 0))
                    Dim TimeTaken As String = ""
                    If SecondsTaken > 60 Then 'Longer than one minute
                        Dim MinutesTaken As Int32 = CInt(Math.Round(SecondsTaken / 60, 0, MidpointRounding.AwayFromZero) - 1)
                        Dim SecondsRemaining = SecondsTaken - MinutesTaken * 60
                        TimeTaken = String.Format(EnglishGB, "{0:0} minutes {1:00} seconds", {MinutesTaken, SecondsRemaining})
                    Else
                        TimeTaken = String.Format(EnglishGB, "{0:0}", SecondsTaken) & " s"
                    End If

                    If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, "File processing complete!", "Total sync directories size: " & SyncSizeString & NewLine & "Time taken: " & TimeTaken, ToolTipIcon.Info)
                Else
                    MyLog.Write("Sync failed: " & Result.ErrorMessage, Warning)
                    If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, "Sync Failed!", Result.ErrorMessage, ToolTipIcon.Error)
                End If
            Else 'Something went badly wrong
                MyLog.Write("Sync failed: no result from background worker.", Warning)
                If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, "Sync Failed!", "No result from background worker", ToolTipIcon.Error)
            End If
        End If

        EnableDisableControls(True)
        ApplyEnableSync()

    End Sub
#End Region

End Class