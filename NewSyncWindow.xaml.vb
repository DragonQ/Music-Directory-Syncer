#Region " Namespaces "
Imports DragonQ.Toolkit
Imports Music_Folder_Syncer.Codec.CodecType
Imports Music_Folder_Syncer.Logger.DebugLogLevel
Imports System.IO
Imports System.Environment
Imports System.Threading
Imports System.Security.AccessControl
Imports System.Security.Principal
Imports Microsoft.WindowsAPICodePack.Dialogs
Imports CodeProject
#End Region


Public Class NewSyncWindow
    Dim AvailableCodecs As New List(Of CheckedListItem)

    Dim WithEvents SyncBackgroundWorker As New BackgroundWorker
    Dim ExitApplication As Boolean = False
    Dim ThreadsCompleted As Int64 = 0
    Dim SyncFolderSize As Int64
    Dim SyncTimer As New Stopwatch()

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        SyncBackgroundWorker.WorkerReportsProgress = True
        SyncBackgroundWorker.WorkerSupportsCancellation = True
        AddHandler SyncBackgroundWorker.DoWork, AddressOf SyncFolder
        AddHandler SyncBackgroundWorker.ProgressChanged, AddressOf SyncFolderProgressChanged
        AddHandler SyncBackgroundWorker.RunWorkerCompleted, AddressOf SyncFolderCompleted

        Dim CodecCount As Int32 = 0

        ' Add all codecs (previously read from Codecs.xml) to AvailableCodecs list
        For Each MyCodec As Codec In Codecs
            Dim CodecString As String = MyCodec.Name & " ("
            For Each MyExtension As String In MyCodec.FileExtensions
                CodecString &= MyExtension & ", "
            Next
            AvailableCodecs.Add(New CheckedListItem(MyCodec, CodecString.Substring(0, CodecString.Length - 2) & ")", False))

            If MyCodec.Type = Lossy Then
                cmbCodec.Items.Add(New Item(MyCodec.Name, MyCodec))
                If MyCodec.Name = DefaultSyncSettings.Encoder.Name Then
                    cmbCodec.SelectedIndex = CodecCount
                End If
                CodecCount += 1
            End If
        Next

        Dim CodecsFilter As Codec() = DefaultSyncSettings.GetWatcherCodecs
        For Each MyFilter As Codec In CodecsFilter
            For Each MyListItem As CheckedListItem In AvailableCodecs
                If MyFilter.Name = CType(MyListItem.Data, Codec).Name Then
                    MyListItem.IsChecked = True
                    Exit For
                End If
            Next
        Next

        For Each MyExtension As CheckedListItem In AvailableCodecs
            lstFileTypesToSync.Items.Add(CType(MyExtension, CheckedListItem))
        Next

        spinThreads.Maximum = DefaultSyncSettings.MaxThreads
        spinThreads.Value = spinThreads.Maximum
        spinThreads.IsEnabled = False
        txt_ffmpegPath.Text = DefaultSyncSettings.ffmpegPath
        txtSourceDirectory.Text = DefaultSyncSettings.SourceDirectory
        txtSyncDirectory.Text = DefaultSyncSettings.SyncDirectory

    End Sub

#Region " Window Controls "
    Private Sub btnBrowseSourceDirectory_Click(sender As Object, e As RoutedEventArgs) Handles btnBrowseSourceDirectory.Click

        Dim SelectDirectoryDialog = New CommonOpenFileDialog()
        SelectDirectoryDialog.Title = "Select Sync Directory"
        SelectDirectoryDialog.IsFolderPicker = True
        SelectDirectoryDialog.AddToMostRecentlyUsedList = False
        SelectDirectoryDialog.AllowNonFileSystemItems = True
        SelectDirectoryDialog.DefaultDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        SelectDirectoryDialog.EnsureFileExists = False
        SelectDirectoryDialog.EnsurePathExists = True
        SelectDirectoryDialog.EnsureReadOnly = False
        SelectDirectoryDialog.EnsureValidNames = True
        SelectDirectoryDialog.Multiselect = False
        SelectDirectoryDialog.ShowPlacesList = True

        If SelectDirectoryDialog.ShowDialog() = CommonFileDialogResult.Ok Then
            txtSourceDirectory.Text = SelectDirectoryDialog.FileName
        End If

    End Sub

    Private Sub btnBrowseSyncDirectory_Click(sender As Object, e As RoutedEventArgs) Handles btnBrowseSyncDirectory.Click

        Dim SelectDirectoryDialog = New CommonOpenFileDialog()
        SelectDirectoryDialog.Title = "Select Sync Directory"
        SelectDirectoryDialog.IsFolderPicker = True
        SelectDirectoryDialog.AddToMostRecentlyUsedList = False
        SelectDirectoryDialog.AllowNonFileSystemItems = True
        SelectDirectoryDialog.DefaultDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        SelectDirectoryDialog.EnsureFileExists = False
        SelectDirectoryDialog.EnsurePathExists = True
        SelectDirectoryDialog.EnsureReadOnly = False
        SelectDirectoryDialog.EnsureValidNames = True
        SelectDirectoryDialog.Multiselect = False
        SelectDirectoryDialog.ShowPlacesList = True

        If SelectDirectoryDialog.ShowDialog() = CommonFileDialogResult.Ok Then
            txtSyncDirectory.Text = SelectDirectoryDialog.FileName
        End If

    End Sub

    Private Sub btnBrowseFFMPEG_Click(sender As Object, e As RoutedEventArgs) Handles btnBrowseFFMPEG.Click

        Dim SelectDirectoryDialog = New CommonOpenFileDialog()
        SelectDirectoryDialog.Title = "Select ffmpeg.exe Path"
        SelectDirectoryDialog.IsFolderPicker = False
        SelectDirectoryDialog.AddToMostRecentlyUsedList = False
        SelectDirectoryDialog.AllowNonFileSystemItems = True
        SelectDirectoryDialog.DefaultDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
        SelectDirectoryDialog.EnsureFileExists = True
        SelectDirectoryDialog.EnsurePathExists = True
        SelectDirectoryDialog.EnsureReadOnly = False
        SelectDirectoryDialog.EnsureValidNames = True
        SelectDirectoryDialog.Multiselect = False
        SelectDirectoryDialog.ShowPlacesList = True
        SelectDirectoryDialog.Filters.Add(New CommonFileDialogFilter("ffmpeg Executable", "*.exe"))

        If SelectDirectoryDialog.ShowDialog() = CommonFileDialogResult.Ok Then
            txt_ffmpegPath.Text = SelectDirectoryDialog.FileName
        End If

    End Sub

    Private Sub cmbCodec_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbCodec.SelectionChanged

        If cmbCodec.SelectedIndex > -1 Then
            cmbCodecLevel.Items.Clear()
            Dim MyCodecLevels As Codec.Profile() = CType(CType(cmbCodec.SelectedItem, Item).Value, Codec).Profiles

            If MyCodecLevels.Count > 0 Then
                For Each CodecLevel As Codec.Profile In MyCodecLevels ' CType(CType(cmbCodec.SelectedItem, Item).Value, Codec).Profiles
                    cmbCodecLevel.Items.Add(New Item(CodecLevel.Name, CodecLevel))
                Next
                cmbCodecLevel.SelectedIndex = 0
            End If

        End If

    End Sub

    Private Sub tckTranscode_Ticked(sender As Object, e As RoutedEventArgs)
        If tckTranscode.IsLoaded AndAlso boxTranscodeOptions.IsLoaded Then
            boxTranscodeOptions.IsEnabled = True
        End If

    End Sub

    Private Sub tckTranscode_Unticked(sender As Object, e As RoutedEventArgs)
        If tckTranscode.IsLoaded AndAlso boxTranscodeOptions.IsLoaded Then
            boxTranscodeOptions.IsEnabled = False
        End If
    End Sub

    Private Sub btnNewSync_Click(sender As Object, e As RoutedEventArgs)

        If SyncBackgroundWorker.IsBusy Then 'Sync in progress, so we must cancel
            If System.Windows.MessageBox.Show("Are you sure you want to cancel this sync operation? Your sync directory will be incomplete!", "Sync in progress!",
                                   MessageBoxButton.OKCancel, MessageBoxImage.Error) = MessageBoxResult.OK Then
                SyncBackgroundWorker.CancelAsync()
            End If
        Else 'Begin sync process
            If System.Windows.MessageBox.Show("Are you sure you want to start a new sync? This will delete all files and folders in the specified " &
                    "sync folder!", "New Sync", MessageBoxButton.OKCancel, MessageBoxImage.Warning) = MessageBoxResult.OK Then
                EnableDisableControls(False)
                MyLog.Write("Creating new folder sync!", Warning)
                MyLog.Write("Source directory: """ & txtSourceDirectory.Text & """.", Information)
                MyLog.Write("Sync directory: """ & txtSyncDirectory.Text & """.", Information)

                MySyncSettings.SourceDirectory = txtSourceDirectory.Text
                MySyncSettings.SyncDirectory = txtSyncDirectory.Text
                MySyncSettings.TranscodeLosslessFiles = tckTranscode.IsChecked
                MySyncSettings.MaxThreads = CInt(spinThreads.Value)
                MySyncSettings.ffmpegPath = txt_ffmpegPath.Text

                If cmbCodec.SelectedIndex > -1 Then
                    If cmbCodecLevel.SelectedIndex > -1 Then
                        MySyncSettings.Encoder = New Codec(CType(CType(cmbCodec.SelectedItem, Item).Value, Codec),
                                                           CType(CType(cmbCodecLevel.SelectedItem, Item).Value, Codec.Profile))
                    End If
                End If

                Dim CodecFilterList As New List(Of Codec)
                For Each TickedCodec As CheckedListItem In AvailableCodecs
                    If TickedCodec.IsChecked Then CodecFilterList.Add(CType(TickedCodec.Data, Codec))
                Next
                MySyncSettings.SetWatcherCodecs(CodecFilterList)

                FilesCompletedProgressBar.Value = 0
                FilesCompletedProgressBar.IsIndeterminate = True
                txtFilesRemaining.Text = "???"
                txtFilesProcessed.Text = "???"
                btnNewSync.Content = "Scanning files..."
                SyncTimer.Start()
                SyncBackgroundWorker.RunWorkerAsync(MySyncSettings.SourceDirectory)
            End If
        End If

    End Sub

    Private Sub EnableDisableControls(Enable As Boolean)
        btnNewSync.IsEnabled = Enable
        tckTranscode.IsEnabled = Enable
        txtSourceDirectory.IsEnabled = Enable
        txtSyncDirectory.IsEnabled = Enable
    End Sub
#End Region

#Region " Start New Sync [Background Worker] "
    Private Sub SyncFolder(sender As Object, e As DoWorkEventArgs)

        Dim FolderPath As String = CType(e.Argument, String)
        Dim MyFiles As FileData() = Nothing
        Dim MyFilesToProcess As New List(Of String)
        Dim CodecsToCheck As Codec() = MySyncSettings.GetWatcherCodecs

        MyLog.Write("Scanning files in source directory.", Information)

        Try
            MyFiles = FastDirectoryEnumerator.EnumerateFiles(FolderPath, "*.*", SearchOption.AllDirectories).ToArray
        Catch ex As Exception
            MyLog.Write("Failed to grab file list from source directory. Exception: " & ex.Message, Warning)
            e.Result = New ReturnObject(False, ex.Message, "")
            Exit Sub
        End Try

        If MyFiles Is Nothing Then
            MyLog.Write("Failed to grab file list from source directory. Exception: EnumerateFiles returned nothing.", Warning)
            e.Result = New ReturnObject(False, "EnumerateFiles returned nothing.", "")
            Exit Sub
        ElseIf MyFiles.Length < 1 Then
            MyLog.Write("Failed to grab file list from source directory. Exception: EnumerateFiles returned no files.", Warning)
            e.Result = New ReturnObject(False, "EnumerateFiles returned no files.", "")
            Exit Sub
        End If

        MyLog.Write("Creating sync folder.", Information)
        Dim ThreadsToRun As Int32 = MyFiles.Count
        Dim ThreadsStarted As UInt32 = 0
        Dim One As UInt32 = 1

        'Delete existing sync folder if necessary
        Try
            Directory.Delete(MySyncSettings.SyncDirectory, True)
        Catch ex As DirectoryNotFoundException
            'Do nothing
        Catch ex As Exception
            MyLog.Write("Failed to delete existing files in sync folder. Exception: " & ex.Message & NewLine &
                     NewLine & "                  " & ex.InnerException.ToString, Warning)
            e.Result = New ReturnObject(False, ex.Message, "")
            Exit Sub
        End Try

        'Create sync folder
        Try
            Dim sid = New SecurityIdentifier(WellKnownSidType.AuthenticatedUserSid, Nothing)
            Dim FullAccess As New DirectorySecurity()
            FullAccess.AddAccessRule(New FileSystemAccessRule(sid, FileSystemRights.FullControl, AccessControlType.Allow))

            Directory.CreateDirectory(MySyncSettings.SyncDirectory, FullAccess)
        Catch ex As Exception
            MyLog.Write("Failed to delete existing files in sync folder. Exception: " & ex.Message & NewLine &
                     NewLine & "                  " & ex.InnerException.ToString, Warning)
            e.Result = New ReturnObject(False, ex.Message, "")
            Exit Sub
        End Try

        MyLog.Write("Starting file processing threads.", Information)
        SyncFolderSize = 0

        Try
            '==============================================================================================
            '============== My own custom way of handling threads, replaced with ThreadPool ===============
            '==============================================================================================
            'For Each MyFile As FileData In MyFiles
            '    Do
            '        Dim MyThreadsRunning As Int64 = Interlocked.Read(ThreadsRunning)
            '        SyncBackgroundWorker.ReportProgress(CInt(Interlocked.Read(ThreadsCompleted) / ThreadsToRun * 100), MyThreadsRunning)

            '        If MyThreadsRunning < MySyncSettings.MaxThreads Then
            '            Interlocked.Increment(ThreadsRunning)

            '            Dim NewThread As New Thread(New ParameterizedThreadStart(AddressOf TransferToSyncFolderDelegate))
            '            NewThread.IsBackground = True
            '            Dim InputObjects As Object() = {MyFile.Path, CodecsToCheck}

            '            NewThread.Start(InputObjects)

            '            ThreadsStarted += One
            '            MyLog.Write("[" & NewThread.ManagedThreadId & "] Processing file: """ & MyFile.Path & """...", Debug)
            '            Exit Do
            '        Else
            '            Thread.Sleep(50)
            '        End If
            '    Loop
            'Next

            ''All threads have been started, now wait for them all to finish
            'Dim CheckPointThreadsCompleted As Int64 = Interlocked.Read(ThreadsCompleted)

            'Do
            '    Dim ThreadsCompletedSoFar As Int64 = Interlocked.Read(ThreadsCompleted)

            '    If ThreadsCompletedSoFar >= ThreadsStarted Then
            '        SyncBackgroundWorker.ReportProgress(100, 0)
            '        MyLog.Write("All files processed. " & ThreadsStarted & " files synced.", Information)
            '        Exit Do
            '    Else
            '        If ThreadsCompletedSoFar > CheckPointThreadsCompleted Then
            '           SyncBackgroundWorker.ReportProgress(CInt(ThreadsCompletedSoFar / ThreadsStarted * 100), Interlocked.Read(ThreadsRunning))
            '        End If
            '        Thread.Sleep(50)
            '    End If
            'Loop

            '==============================================================================================
            '==============================================================================================

            ThreadPool.SetMinThreads(MySyncSettings.MaxThreads, MySyncSettings.MaxThreads)
            ThreadPool.SetMaxThreads(MySyncSettings.MaxThreads, MySyncSettings.MaxThreads)

            For Each MyFile As FileData In MyFiles
                ThreadsStarted += One
                Dim InputObjects As Object() = {ThreadsStarted, MyFile.Path, CodecsToCheck}
                ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf TransferToSyncFolderDelegate), InputObjects)
            Next

            'All threads have been started, now wait for them all to finish
            Dim CheckPointThreadsCompleted As Int64 = Interlocked.Read(ThreadsCompleted)
            Dim ThreadsRemaining As Int64 = ThreadsStarted - CheckPointThreadsCompleted
            SyncBackgroundWorker.ReportProgress(CInt(CheckPointThreadsCompleted / ThreadsStarted * 100),
                                                            {ThreadsRemaining, CheckPointThreadsCompleted})

            Do
                If SyncBackgroundWorker.CancellationPending Then
                    e.Cancel = True
                    e.Result = New ReturnObject(False, "Sync was cancelled.", "")
                    Exit Sub
                End If

                Thread.Sleep(500) 'Update UI twice a second

                Dim ThreadsCompletedSoFar As Int64 = Interlocked.Read(ThreadsCompleted)

                If ThreadsCompletedSoFar > CheckPointThreadsCompleted Then 'At least one thread has finished since the last check
                    If ThreadsCompletedSoFar >= ThreadsStarted Then 'All files have been processed, so end background worker
                        MyLog.Write("All files processed. " & ThreadsStarted & " files synced.", Information)
                        Exit Do
                    Else
                        ThreadsRemaining = ThreadsStarted - ThreadsCompletedSoFar
                        SyncBackgroundWorker.ReportProgress(CInt(ThreadsCompletedSoFar / ThreadsStarted * 100),
                                                            {ThreadsRemaining, ThreadsCompletedSoFar})
                    End If
                End If
            Loop
        Catch ex As Exception
            MyLog.Write("Failed to complete sync. Exception: " & ex.Message, Warning)
            e.Result = New ReturnObject(False, ex.Message, "")
        End Try

        e.Result = New ReturnObject(True, Nothing, SyncFolderSize)

    End Sub

    Private Sub TransferToSyncFolderDelegate(ByVal Input As Object)

        Try
            Dim InputObjects As Object() = CType(Input, Object())
            Dim ProcessID As UInt32 = CType(InputObjects(0), UInt32)
            Dim TransferResult As ReturnObject

            Try
                TransferResult = TransferToSyncFolder(ProcessID, CType(InputObjects(1), String), CType(InputObjects(2), Codec()))

                If TransferResult.Success Then
                    Dim NewSize As Int64 = CType(TransferResult.MyObject, Int64)

                    If NewSize > 0 Then
                        Interlocked.Add(SyncFolderSize, NewSize)
                    End If
                Else
                    MyLog.Write(ProcessID, "File could not be processed. Error: " & TransferResult.ErrorMessage, Warning)
                End If
            Catch ex As Exception
                MyLog.Write(ProcessID, "File could not be processed. Malformed input to TransferToSyncFolderDelegate subroutine.", Warning)
            End Try
        Catch ex As Exception
            MyLog.Write(UInt32.MaxValue, "File could not be processed. Malformed input to TransferToSyncFolderDelegate.", Warning)
        End Try

        Interlocked.Increment(ThreadsCompleted)
        'Interlocked.Decrement(ThreadsRunning)

    End Sub

    Private Sub SyncFolderCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs)

        SyncTimer.Stop()
        txtFilesRemaining.Text = ""
        txtFilesProcessed.Text = ""
        FilesCompletedProgressBar.Value = 100

        'If the sync task was cancelled, then we need to force-close all instances of ffmpeg and exit application
        If e.Cancelled Then
            Try
                Dim taskkill As New ProcessStartInfo("taskkill")
                taskkill.CreateNoWindow = True
                taskkill.UseShellExecute = False
                taskkill.Arguments = " /F /IM " & Path.GetFileName(MySyncSettings.ffmpegPath) & " /T"

                MyLog.Write("Sync cancelled, now force-closing ffmpeg instances...", Warning)

                Dim taskkillProcess As Process = Process.Start(taskkill)
                taskkillProcess.WaitForExit()

                MyLog.Write("...done!", Warning)
            Catch ex As Exception
                MyLog.Write("...failed! There may be lingering ffmpeg instances.", Warning)
            End Try

            If ExitApplication Then
                ExitApplication = False
                Me.Close()
            End If

            Exit Sub
        End If

        Dim Result As ReturnObject = CType(e.Result, ReturnObject)

        If Result.Success Then
            'Work out size of sync folder
            Dim SyncSize As Double = CType(Result.MyObject, Double) / (2 ^ 20) ' Convert to MiB
            Dim SyncSizeString As String = ""
            If SyncSize > 1024 Then ' Directory size is greater than 1 GiB
                SyncSizeString = String.Format("{0:0.0}", SyncSize / (2 ^ 10)) & " GiB"
            Else
                SyncSizeString = String.Format("{0:0.0}", SyncSize) & " MiB"
            End If

            'Work out how long the sync took
            Dim SecondsTaken As Int64 = CInt(Math.Round(SyncTimer.ElapsedMilliseconds / 1000, 0))
            Dim TimeTaken As String = ""
            If SecondsTaken > 60 Then 'Longer than one minute
                Dim MinutesTaken As Int32 = CInt(Math.Round(SecondsTaken / 60, 0, MidpointRounding.AwayFromZero) - 1)
                Dim SecondsRemaining = SecondsTaken - MinutesTaken * 60
                TimeTaken = String.Format("{0:0} minutes {1:00} seconds", {MinutesTaken, SecondsRemaining})
            Else
                TimeTaken = String.Format("{0:0}", SecondsTaken) & " s"
            End If

            'Ask user if they want to start the background sync updater (presumably yes)
            If System.Windows.MessageBox.Show("Sync directory size: " & SyncSizeString & NewLine & NewLine & "Time taken: " & TimeTaken & NewLine & NewLine &
                               "Do you want to enable background sync " &
                                    "for this folder? This will ensure your sync folder is always up-to-date.", "Sync Complete!",
                                    MessageBoxButton.OKCancel, MessageBoxImage.Information) = MessageBoxResult.OK Then
                MySyncSettings.SyncIsEnabled = True
            Else
                MySyncSettings.SyncIsEnabled = False
            End If

            Dim MyResult As ReturnObject = SaveSyncSettings()

            If MyResult.Success Then
                Me.DialogResult = True
                MyLog.Write("Syncer settings updated.", Information)
                Me.Close()
            Else
                MyLog.Write("Could not update sync settings. Error: " & MyResult.ErrorMessage, Warning)
                System.Windows.MessageBox.Show("Could not save sync settings!", MyResult.ErrorMessage, MessageBoxButton.OK, MessageBoxImage.Error)
            End If
        Else
            System.Windows.MessageBox.Show("Sync failed! " & NewLine & NewLine & Result.ErrorMessage, "Sync Failed!",
                    MessageBoxButton.OKCancel, MessageBoxImage.Error)
            EnableDisableControls(True)
        End If

    End Sub

    Private Sub SyncFolderProgressChanged(sender As Object, e As ProgressChangedEventArgs)

        'If Not e.UserState Is Nothing Then txtThreadsRunning.Text = "Threads running: " & CType(e.UserState, Int64).ToString
        If Not e.UserState Is Nothing Then
            Dim Times As Int64() = CType(e.UserState, Int64())
            txtFilesRemaining.Text = Times(0).ToString
            txtFilesProcessed.Text = Times(1).ToString
        End If
        FilesCompletedProgressBar.IsIndeterminate = False
        btnNewSync.Content = "Processing files..."
        FilesCompletedProgressBar.Value = e.ProgressPercentage

    End Sub
#End Region

#Region " Window Closing "
    Private Sub MainWindow_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        'User has closed the program, so we need to check if a sync is occuring
        If SyncBackgroundWorker.IsBusy Then
            If System.Windows.MessageBox.Show("Are you sure you want to exit? Your sync directory will be incomplete!", "Sync in progress!",
                               MessageBoxButton.OKCancel, MessageBoxImage.Error) = MessageBoxResult.OK Then
                SyncBackgroundWorker.CancelAsync()
                ExitApplication = True
            End If
            e.Cancel = True
        ElseIf ExitApplication Then
            e.Cancel = True
        End If

        If DialogResult Is Nothing Then
            DialogResult = False
        End If

    End Sub
#End Region

End Class