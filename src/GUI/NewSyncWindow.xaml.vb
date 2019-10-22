#Region " Namespaces "
Imports MusicFolderSyncer.Toolkit
Imports MusicFolderSyncer.Toolkit.Browsers
Imports MusicFolderSyncer.Codec.CodecType
Imports MusicFolderSyncer.Logger.LogLevel
Imports MusicFolderSyncer.SyncSettings
Imports MusicFolderSyncer.SyncSettings.TranscodeMode
Imports System.IO
Imports System.Environment
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
#End Region

Public Class NewSyncWindow

    Private MySyncer As SyncerInitialiser = Nothing
    Private ExitApplication As Boolean = False
    Private ReadOnly SyncTimer As New Stopwatch()
    Private TagsToSync As ObservableCollection(Of Codec.Tag)
    Private FileTypesToSync As ObservableCollection(Of Codec)
    Private MyGlobalSyncSettings As GlobalSyncSettings
    Private MySyncSettings As SyncSettings
    Private MySyncSettingsList As List(Of SyncSettings)
    Private IsNewSync As Boolean
    Private SyncInProgress As Boolean = False

#Region " New "
    Public Sub New(NewGlobalSyncSettings As GlobalSyncSettings, MyNewSync As Boolean)

        ' This call is required by the designer.
        InitializeComponent()
        Me.DataContext = Me

        ' Define default sync settings to start with
        MyGlobalSyncSettings = NewGlobalSyncSettings
        MySyncSettingsList = MyGlobalSyncSettings.GetSyncSettings().ToList()
        MySyncSettings = New SyncSettings(MySyncSettingsList(0))
        IsNewSync = MyNewSync

        ' Define collection for FileTypesToSync, which is bound to lstFileTypesToSync
        FileTypesToSync = New ObservableCollection(Of Codec)
        AddHandler FileTypesToSync.CollectionChanged, AddressOf FileTypesToSyncResized
        Dim CodecsFilter As Codec() = MySyncSettings.GetWatcherCodecs
        For Each MyFilter As Codec In CodecsFilter
            FileTypesToSync.Add(MyFilter)
        Next

        ' Define collection for TagsToSync, which is bound to dataTagsToSync
        TagsToSync = New ObservableCollection(Of Codec.Tag)
        Dim Tags As Codec.Tag() = MySyncSettings.GetWatcherTags
        For Each MyTag As Codec.Tag In Tags
            TagsToSync.Add(MyTag)
        Next

    End Sub

    Private Sub NewSyncWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

        ' Add all codecs (previously read from Codecs.xml) to cmbCodec
        Dim Count As Int32 = 0
        For Each MyCodec As Codec In Codecs
            If MyCodec.CompressionType = Lossy Then
                cmbCodec.Items.Add(New Item(MyCodec.Name, MyCodec))
                If MyCodec.Name = MySyncSettings.Encoder.Name Then
                    cmbCodec.SelectedIndex = Count
                End If
                Count += 1
            End If
        Next

        ' Add all ReplayGain settings to cmbReplayGain
        Dim Enums As Array = System.Enum.GetValues(GetType(ReplayGainMode))
        Count = 0
        For Each MyEnum In Enums
            Dim ReplayGainSetting As ReplayGainMode = DirectCast(MyEnum, ReplayGainMode)
            Dim Name As String = SyncSettings.GetReplayGainSetting(ReplayGainSetting)
            cmbReplayGain.Items.Add(New Item(Name, ReplayGainSetting))
            If ReplayGainSetting = MySyncSettings.ReplayGainSetting Then
                cmbReplayGain.SelectedIndex = Count
            End If
            Count += 1
        Next

        ' Add all transcoding settings to cmbTranscodeSetting
        Enums = System.Enum.GetValues(GetType(TranscodeMode))
        Count = 0
        For Each MyEnum In Enums
            Dim TranscodeSetting As TranscodeMode = DirectCast(MyEnum, TranscodeMode)
            Dim Name As String = SyncSettings.GetTranscodeSetting(TranscodeSetting)
            cmbTranscodeSetting.Items.Add(New Item(Name, TranscodeSetting))
            If TranscodeSetting = MySyncSettings.TranscodeSetting Then
                cmbTranscodeSetting.SelectedIndex = Count
            End If
            Count += 1
        Next

        ' Set run-time properties of window objects
        If TagsToSync.Count > 0 Then btnRemoveTag.IsEnabled = True
        spinThreads.Maximum = MyGlobalSyncSettings.MaxThreads
        spinThreads.Value = spinThreads.Maximum
        txt_ffmpegPath.Text = MyGlobalSyncSettings.ffmpegPath
        txtSourceDirectory.Text = MyGlobalSyncSettings.SourceDirectory
        txtSyncDirectory.Text = MySyncSettings.SyncDirectory
        txtSourceDirectory.Focus()

        ' Disable fields if this is an additional sync setup
        If Not IsNewSync Then
            txt_ffmpegPath.IsEnabled = False
            btnBrowse_ffmpeg.IsEnabled = False
            txtSourceDirectory.IsEnabled = False
            btnBrowseSourceDirectory.IsEnabled = False
        End If

    End Sub
#End Region

    Public ReadOnly Property GetTagsToSync() As ObservableCollection(Of Codec.Tag)
        Get
            If TagsToSync Is Nothing Then
                TagsToSync = New ObservableCollection(Of Codec.Tag)()
            End If
            Return TagsToSync
        End Get
    End Property

    Public ReadOnly Property GetFileTypesToSync() As ObservableCollection(Of Codec)
        Get
            Dim EnabledFileTypesToSync As New ObservableCollection(Of Codec)

            ' Return a new list of codecs that only includes the codecs enabled/ticked by the user
            If Not FileTypesToSync Is Nothing Then
                For Each FileType As Codec In FileTypesToSync
                    If FileType.IsEnabled Then EnabledFileTypesToSync.Add(FileType)
                Next
            End If

            Return EnabledFileTypesToSync
        End Get
    End Property

    Private Sub FileTypesToSyncResized(sender As Object, e As NotifyCollectionChangedEventArgs)

        If e.Action = NotifyCollectionChangedAction.Add Then
            ' Check whether the user has chosen to monitor WMA files and enable/disable the special tick box as appropriate.
            For Each FileType As Codec In e.NewItems
                If CheckForWMA(FileType) = True Then Exit For
            Next
        End If

    End Sub

    Private Sub FileTypesToSyncChanged(sender As Object, e As RoutedEventArgs)

        ' Check whether the user has chosen to monitor WMA files and enable/disable the special tick box as appropriate.
        For Each FileType As Codec In FileTypesToSync
            If CheckForWMA(FileType) = True Then Exit For
        Next

    End Sub

    Private Function CheckForWMA(FileType As Codec) As Boolean
        If FileType.Name = "WMA" Then
            If FileType.IsEnabled Then
                tckTreatWMA_AsLossless.IsEnabled = True
                FileType.SetType(Lossless)
            Else
                tckTreatWMA_AsLossless.IsEnabled = False
                FileType.SetType(Lossless)
            End If
            Return True
        Else
            Return False
        End If
    End Function

#Region " Window Controls "
    Private Sub btnNewTag_Click(sender As Object, e As RoutedEventArgs)
        TagsToSync.Add(New Codec.Tag("Tag Name", "Tag Value"))
        btnRemoveTag.IsEnabled = True
    End Sub

    Private Sub btnRemoveTag_Click(sender As Object, e As RoutedEventArgs)

        If dataTagsToSync.SelectedIndex >= 0 Then
            TagsToSync.RemoveAt(dataTagsToSync.SelectedIndex)
            If TagsToSync.Count <= 0 Then btnRemoveTag.IsEnabled = False
        End If

    End Sub

    Private Sub btnBrowseSourceDirectory_Click(sender As Object, e As RoutedEventArgs)

        Dim DefaultDirectory As String = txtSourceDirectory.Text

        If Not Directory.Exists(DefaultDirectory) Then
            DefaultDirectory = MyGlobalSyncSettings.SourceDirectory
        End If

        Dim Browser As ReturnObject = CreateDirectoryBrowser(DefaultDirectory, "Select Source Directory")

        If Browser.Success Then
            txtSourceDirectory.Text = CStr(Browser.MyObject)
        End If

    End Sub

    Private Sub btnBrowseSyncDirectory_Click(sender As Object, e As RoutedEventArgs)

        Dim DefaultDirectory As String = txtSyncDirectory.Text

        If Not Directory.Exists(DefaultDirectory) Then
            DefaultDirectory = MySyncSettings.SyncDirectory
        End If

        Dim Browser As ReturnObject = CreateDirectoryBrowser(DefaultDirectory, "Select Sync Directory")

        If Browser.Success Then
            txtSyncDirectory.Text = CStr(Browser.MyObject)
        End If

    End Sub

    Private Sub btnBrowseFFMPEG_Click(sender As Object, e As RoutedEventArgs)

        Dim DefaultPath As String = txt_ffmpegPath.Text

        If Not Directory.Exists(DefaultPath) Then
            DefaultPath = MyGlobalSyncSettings.ffmpegPath
        End If

        Dim Browser As ReturnObject = CreateFileBrowser_ffmpeg(DefaultPath)

        If Browser.Success Then
            txt_ffmpegPath.Text = CStr(Browser.MyObject)
        End If

    End Sub

    Private Sub cmbCodec_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbCodec.SelectionChanged

        If cmbCodec.SelectedIndex > -1 Then
            cmbCodecProfile.Items.Clear()
            Dim MyCodecLevels As Codec.Profile() = CType(CType(cmbCodec.SelectedItem, Item).Value, Codec).GetProfiles()

            If MyCodecLevels.Count > 0 Then
                For Each CodecLevel As Codec.Profile In MyCodecLevels ' CType(CType(cmbCodec.SelectedItem, Item).Value, Codec).Profiles
                    cmbCodecProfile.Items.Add(New Item(CodecLevel.Name, CodecLevel))
                Next
                cmbCodecProfile.SelectedIndex = 0
            End If

        End If

    End Sub

    Private Sub cmbTranscodeSetting_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbTranscodeSetting.SelectionChanged

        If cmbTranscodeSetting.SelectedIndex > -1 Then
            MySyncSettings.TranscodeSetting = CType(CType(cmbTranscodeSetting.SelectedItem, Item).Value, TranscodeMode)
            If MySyncSettings.TranscodeSetting = None Then
                boxTranscodeOptions.IsEnabled = False
            Else
                boxTranscodeOptions.IsEnabled = True
            End If
        Else
            System.Windows.MessageBox.Show("You have not specified a valid transcode setting for this sync!", "New Sync", MessageBoxButton.OKCancel, MessageBoxImage.Error)
        End If

    End Sub

    Private Sub CancelSync()
        If Not MySyncer Is Nothing Then
            MySyncer.SyncBackgroundWorker.CancelAsync()
        End If
    End Sub

    Private Sub btnNewSync_Click(sender As Object, e As RoutedEventArgs)

        If SyncInProgress Then 'Sync in progress, so we must cancel
            If System.Windows.MessageBox.Show("Are you sure you want to cancel this sync operation? Your sync directory will be incomplete!", "Sync in progress!",
                                   MessageBoxButton.OKCancel, MessageBoxImage.Error) = MessageBoxResult.OK Then
                CancelSync()
            End If
        Else 'Begin sync process
            If System.Windows.MessageBox.Show("Are you sure you want to start a new sync? This will delete all files and folders in the specified " &
                    "sync folder!", "New Sync", MessageBoxButton.OKCancel, MessageBoxImage.Warning) = MessageBoxResult.OK Then

                'Disable window controls
                EnableDisableControls(False)

                'Grab list of files to sync and check if there are any listed. If not, abort sync creation.
                Dim NewFileTypesToSync As List(Of Codec) = GetFileTypesToSync().ToList
                If NewFileTypesToSync Is Nothing OrElse NewFileTypesToSync.Count = 0 Then
                    System.Windows.MessageBox.Show("You have not specified any file types to match!", "New Sync", MessageBoxButton.OKCancel, MessageBoxImage.Error)
                    EnableDisableControls(True)
                    Exit Sub
                End If

                'Grab list of tags to sync and check if there are any listed. If not, show a warning.
                Dim NewTagsToSync As List(Of Codec.Tag) = GetTagsToSync().ToList
                If NewTagsToSync Is Nothing OrElse NewTagsToSync.Count = 0 Then
                    If System.Windows.MessageBox.Show("You have not specified any tags to match. All files with specified file types will be synced. " & NewLine & NewLine &
                                                  "Are you sure this is what you want to do?", "New Sync", MessageBoxButton.OKCancel, MessageBoxImage.Warning) <> MessageBoxResult.OK Then
                        EnableDisableControls(True)
                        Exit Sub
                    End If
                End If

                MyLog.Write("Creating new folder sync!", Warning)
                MyLog.Write("Source directory: """ & txtSourceDirectory.Text & """.", Information)
                MyLog.Write("Sync directory: """ & txtSyncDirectory.Text & """.", Information)

                'Set transcode setting
                If cmbTranscodeSetting.SelectedIndex > -1 Then
                    MySyncSettings.TranscodeSetting = CType(CType(cmbTranscodeSetting.SelectedItem, Item).Value, TranscodeMode)
                Else
                    System.Windows.MessageBox.Show("You have not specified a valid transcode setting for this sync!", "New Sync", MessageBoxButton.OKCancel, MessageBoxImage.Error)
                    EnableDisableControls(True)
                    Exit Sub
                End If

                'If transcoding is enabled, check that a valid encoder and encoder profile have been selected. If not, abort sync creation.
                If MySyncSettings.TranscodeSetting <> None Then
                    If (cmbCodec.SelectedIndex > -1 AndAlso cmbCodecProfile.SelectedIndex > -1) Then
                        MySyncSettings.Encoder = New Codec(CType(CType(cmbCodec.SelectedItem, Item).Value, Codec),
                                                        CType(CType(cmbCodecProfile.SelectedItem, Item).Value, Codec.Profile))
                    Else
                        System.Windows.MessageBox.Show("You have not specified a valid encoder for this sync!", "New Sync", MessageBoxButton.OKCancel, MessageBoxImage.Error)
                        EnableDisableControls(True)
                        Exit Sub
                    End If
                End If

                'Set ReplayGain setting
                If cmbReplayGain.SelectedIndex > -1 Then
                    MySyncSettings.ReplayGainSetting = CType(CType(cmbReplayGain.SelectedItem, Item).Value, ReplayGainMode)
                Else
                    System.Windows.MessageBox.Show("You have not specified a valid ReplayGain setting for this sync!", "New Sync", MessageBoxButton.OKCancel, MessageBoxImage.Error)
                    EnableDisableControls(True)
                    Exit Sub
                End If

                'Change remaining sync settings as specified by the user
                MyGlobalSyncSettings.SourceDirectory = txtSourceDirectory.Text.TrimEnd("\"c)
                MySyncSettings.SyncDirectory = txtSyncDirectory.Text.TrimEnd("\"c)
                MyGlobalSyncSettings.MaxThreads = CInt(spinThreads.Value)
                MyGlobalSyncSettings.ffmpegPath = txt_ffmpegPath.Text
                MySyncSettings.SetWatcherTags(NewTagsToSync)
                MySyncSettings.SetWatcherCodecs(NewFileTypesToSync)

                'Update GlobalSyncSettings object
                If IsNewSync Then
                    MySyncSettingsList = New List(Of SyncSettings)
                End If
                MySyncSettingsList.Add(MySyncSettings)
                MyGlobalSyncSettings.SetSyncSettings(MySyncSettingsList)

                'Begin sync process.
                FilesCompletedProgressBar.Value = 0
                FilesCompletedProgressBar.IsIndeterminate = True
                txtFilesRemaining.Text = "???"
                txtFilesProcessed.Text = "???"
                btnNewSync.Content = "Scanning files..."
                SyncTimer.Start()
                SyncFolder()
            End If
        End If

    End Sub

    Private Sub EnableDisableControls(Enable As Boolean)

        'Group box controls
        If MySyncSettings.TranscodeSetting <> None Then
            boxTranscodeOptions.IsEnabled = Enable
        End If
        boxFileTypes.IsEnabled = Enable
        boxTags.IsEnabled = Enable

        'Directory controls
        txtSyncDirectory.IsEnabled = Enable
        btnBrowseSyncDirectory.IsEnabled = Enable
        If IsNewSync Then
            txtSourceDirectory.IsEnabled = Enable
            btnBrowseSourceDirectory.IsEnabled = Enable
            spinThreads.IsEnabled = Enable
        End If

        'Miscellaneous controls
        btnNewSync.IsEnabled = Enable
        cmbTranscodeSetting.IsEnabled = Enable
        cmbReplayGain.IsEnabled = Enable

    End Sub
#End Region

#Region " Start New Sync [Background Worker] "
    Private Sub SyncFolder()

        'Create syncer initialiser
        MySyncer = New SyncerInitialiser(MyGlobalSyncSettings, MySyncSettings, 500)

        'Set callback functions for the SyncBackgroundWorker
        MySyncer.AddProgressCallback(AddressOf SyncFolderProgressChanged)
        MySyncer.AddCompletionCallback(AddressOf SyncFolderCompleted)

        'Start sync
        SyncInProgress = True
        MySyncer.InitialiseSync()

    End Sub

    Private Sub SyncFolderProgressChanged(sender As Object, e As ProgressChangedEventArgs)

        'If Not e.UserState Is Nothing Then txtThreadsRunning.Text = "Threads running: " & CType(e.UserState, Int64).ToString
        If Not e.UserState Is Nothing Then
            Dim Times As Int64() = CType(e.UserState, Int64())
            txtFilesRemaining.Text = Times(0).ToString(EnglishGB)
            txtFilesProcessed.Text = Times(1).ToString(EnglishGB)
        End If
        FilesCompletedProgressBar.IsIndeterminate = False
        btnNewSync.Content = "Processing files..."
        FilesCompletedProgressBar.Value = e.ProgressPercentage

    End Sub

    Private Sub SyncFolderCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs)

        SyncTimer.Stop()
        txtFilesRemaining.Text = ""
        txtFilesProcessed.Text = ""
        btnNewSync.Content = "Start New Sync"
        FilesCompletedProgressBar.Value = 100

        SyncInProgress = False

        'If the sync task was cancelled, then we need to force-close all instances of ffmpeg and exit application
        If e.Cancelled Then
            Try
                MyLog.Write("Sync cancelled, now force-closing ffmpeg instances...", Warning)

                Dim taskkill As New ProcessStartInfo("taskkill") With {
                    .CreateNoWindow = True,
                    .UseShellExecute = False,
                    .Arguments = " /F /IM " & Path.GetFileName(MyGlobalSyncSettings.ffmpegPath) & " /T"
                }
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

        If e.Result Is Nothing Then
            MyLog.Write("Sync failed: no result from background worker.", Fatal)
            System.Windows.MessageBox.Show("No result from background worker.", "Sync Failed!", MessageBoxButton.OK, MessageBoxImage.Error)
            FilesCompletedProgressBar.IsIndeterminate = False
            EnableDisableControls(True)
            Exit Sub
        End If

        Dim Result As ReturnObject = CType(e.Result, ReturnObject)

        If Result.Success Then
            'Work out size of sync folder
            Dim SyncSize As Double = CType(Result.MyObject, Double) / (2 ^ 20) ' Convert to MiB
            Dim SyncSizeString As String
            If SyncSize > 1024 Then ' Directory size is greater than 1 GiB
                SyncSizeString = String.Format(EnglishGB, "{0:0.0}", SyncSize / (2 ^ 10)) & " GiB"
            Else
                SyncSizeString = String.Format(EnglishGB, "{0:0.0}", SyncSize) & " MiB"
            End If

            'Work out how long the sync took
            Dim SecondsTaken As Int64 = CInt(Math.Round(SyncTimer.ElapsedMilliseconds / 1000, 0))
            Dim TimeTaken As String
            If SecondsTaken > 60 Then 'Longer than one minute
                Dim MinutesTaken As Int32 = CInt(Math.Round(SecondsTaken / 60, 0, MidpointRounding.AwayFromZero) - 1)
                Dim SecondsRemaining = SecondsTaken - MinutesTaken * 60
                TimeTaken = String.Format(EnglishGB, "{0:0} minutes {1:00} seconds", MinutesTaken, SecondsRemaining)
            Else
                TimeTaken = String.Format(EnglishGB, "{0:0}", SecondsTaken) & " s"
            End If

            'Ask user if they want to start the background sync updater (presumably yes)
            If System.Windows.MessageBox.Show("Sync directory size: " & SyncSizeString & NewLine & NewLine & "Time taken: " & TimeTaken & NewLine & NewLine &
                               "Do you want to enable background sync? This will ensure your sync folders are always up-to-date.", "Sync Complete!",
                                    MessageBoxButton.OKCancel, MessageBoxImage.Information) = MessageBoxResult.OK Then
                MyGlobalSyncSettings.SyncIsEnabled = True
            Else
                MyGlobalSyncSettings.SyncIsEnabled = False
            End If

            Dim MyResult As ReturnObject = SaveSyncSettings(MyGlobalSyncSettings)

            If MyResult.Success Then
                'Set UserGlobalSyncSettings to our newly updated version now that it's been saved
                UserGlobalSyncSettings = MyGlobalSyncSettings
                Me.DialogResult = True
                MyLog.Write("Syncer settings updated.", Information)
                Me.Close()
            Else
                MyLog.Write("Could not update sync settings. Error: " & MyResult.ErrorMessage, Warning)
                System.Windows.MessageBox.Show("Could not save sync settings!", MyResult.ErrorMessage, MessageBoxButton.OK, MessageBoxImage.Error)
            End If
        Else
            MyLog.Write("Sync failed: " & Result.ErrorMessage, Warning)
            System.Windows.MessageBox.Show("Sync failed! " & NewLine & NewLine & Result.ErrorMessage, "Sync Failed!",
                    MessageBoxButton.OK, MessageBoxImage.Error)
            FilesCompletedProgressBar.IsIndeterminate = False
            EnableDisableControls(True)
        End If

    End Sub
#End Region

#Region " Window Closing "
    Private Sub MainWindow_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        'User has closed the program, so we need to check if a sync is occuring
        If SyncInProgress Then
            If System.Windows.MessageBox.Show("Are you sure you want to exit? Your sync directory will be incomplete!", "Sync in progress!",
                               MessageBoxButton.OKCancel, MessageBoxImage.Error) = MessageBoxResult.OK Then
                CancelSync()
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