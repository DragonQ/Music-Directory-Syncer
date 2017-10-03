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
    Private WithEvents Tray As NotifyIcon
    Private WithEvents MainMenu As ContextMenuStrip
    Private WithEvents mnuEditSyncSettings, mnuEnableSync, mnuExit, mnuNewSync, mnuStatus, mnuViewLogFile As ToolStripMenuItem
    Private WithEvents mnuSep1, mnuSep2, mnuSep3 As ToolStripSeparator
    Private Const BalloonTime As Int32 = 8

    Private WithEvents DirectoryWatcher As FileSystemWatcher
    Private WithEvents FileWatcher As FileSystemWatcher
    Private FileID As Int32 = 0
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
        mnuNewSync = New ToolStripMenuItem("Create new sync")
        mnuSep3 = New ToolStripSeparator()
        mnuExit = New ToolStripMenuItem("Exit")
        mnuStatus.Enabled = False
        mnuEnableSync.Enabled = False
        mnuEditSyncSettings.Enabled = False
        MainMenu = New ContextMenuStrip
        MainMenu.Items.AddRange(New ToolStripItem() {mnuStatus, mnuSep1, mnuEnableSync, mnuEditSyncSettings, mnuViewLogFile, mnuSep2, mnuNewSync, mnuSep3, mnuExit})

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
            mnuEditSyncSettings.Enabled = True
            mnuEnableSync.Enabled = True
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

        If UserGlobalSyncSettings.SyncIsEnabled Then 'Disable sync
            Dim WatcherStopResult As ReturnObject = StopWatcher()
            If Not WatcherStopResult.Success Then
                Dim ErrorMessage As String = "File system watcher could not be stopped. " & NewLine & NewLine & WatcherStopResult.ErrorMessage
                MyLog.Write(ErrorMessage, Warning)
                If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, ApplicationName, ErrorMessage, ToolTipIcon.Error)
            End If
        Else 'Enable sync
            Dim WatcherStartResult As ReturnObject = StartWatcher()
            If Not WatcherStartResult.Success Then
                Dim ErrorMessage As String = "File system watcher could not be started. " & NewLine & NewLine & WatcherStartResult.ErrorMessage
                MyLog.Write(ErrorMessage, Warning)
                If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, ApplicationName, ErrorMessage, ToolTipIcon.Error)
            End If
        End If

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
            ' Sync was successfully set up


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
            mnuEditSyncSettings.Enabled = True
            mnuEnableSync.Enabled = True
        Else ' User closed the window before sync was completed
            mnuStatus.Text = "Syncer is not active"
            Tray.ShowBalloonTip(BalloonTime, ApplicationName, "Syncer not set up.", ToolTipIcon.Info)
        End If

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

            UserGlobalSyncSettings.SyncIsEnabled = True
            mnuStatus.Text = "Syncer is active"
            mnuEnableSync.Text = "Disable sync"
            MyLog.Write("File system watcher started (monitoring directory """ & UserGlobalSyncSettings.SourceDirectory & """ for audio files)", Information)
            If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, ApplicationName, "Syncer active.", ToolTipIcon.Info)

            Return SaveSyncSettings(UserGlobalSyncSettings)
        Catch ex As Exception
            Return New ReturnObject(False, ex.Message, Nothing)
        End Try

    End Function

    Private Function StopWatcher() As ReturnObject

        If FileWatcher IsNot Nothing Then FileWatcher.Dispose()

        UserGlobalSyncSettings.SyncIsEnabled = False
        mnuStatus.Text = "Syncer is not active"
        mnuEnableSync.Text = "Enable sync"
        MyLog.Write("File system watcher stopped.", Information)
        If Tray.Visible Then Tray.ShowBalloonTip(BalloonTime, ApplicationName, "Syncer disabled.", ToolTipIcon.Info)

        Return SaveSyncSettings(UserGlobalSyncSettings)

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
        If (sender Is FileWatcher) AndAlso (Not Directory.Exists(e.FullPath)) AndAlso (FilterMatch(e.Name)) Then
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

End Class