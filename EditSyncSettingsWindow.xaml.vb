#Region " Namespaces "
Imports MusicFolderSyncer.Logger.DebugLogLevel
Imports MusicFolderSyncer.Toolkit
Imports System.IO
#End Region

Public Class EditSyncSettingsWindow

    Dim MySyncSettings As SyncSettings

#Region " New "
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

    End Sub

    Private Sub EditSyncSettingsWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

        ' Set run-time properties of window objects
        txtSourceDirectory.Text = MyGlobalSyncSettings.SourceDirectory
        txt_ffmpegPath.Text = MySyncSettings.SyncDirectory
        spinThreads.Maximum = Environment.ProcessorCount
        spinThreads.Value = MySyncSettings.MaxThreads

    End Sub
#End Region

#Region " Window Controls "
    Private Sub btnBrowseSourceDirectory_Click(sender As Object, e As RoutedEventArgs)

        Dim DefaultDirectory As String = txtSourceDirectory.Text

        If Not Directory.Exists(DefaultDirectory) Then
            DefaultDirectory = MyGlobalSyncSettings.SourceDirectory
        End If

        Dim Browser As ReturnObject = CreateDirectoryBrowser(DefaultDirectory)

        If Browser.Success Then
            txtSourceDirectory.Text = CStr(Browser.MyObject)
        End If

    End Sub

    Private Sub btnBrowse_ffmmpegPath_Click(sender As Object, e As RoutedEventArgs)

        Dim DefaultPath As String = txt_ffmpegPath.Text

        If Not Directory.Exists(DefaultPath) Then
            DefaultPath = MyGlobalSyncSettings.ffmpegPath
        End If

        Dim Browser As ReturnObject = CreateFileBrowser(DefaultPath)

        If Browser.Success Then
            txt_ffmpegPath.Text = CStr(Browser.MyObject)
        End If

    End Sub

    Private Sub btnSave_Click(sender As Object, e As RoutedEventArgs)

        EnableDisableControls(False)

        Try
            ' Check settings aren't nonsense
            If Not Directory.Exists(txtSourceDirectory.Text) Then
                Throw New Exception("Specified source directory doesn't exist.")
            End If

            If Not File.Exists(txt_ffmpegPath.Text) Then
                Throw New Exception("Specified ffmpeg path is not valid.")
            End If

            ' Apply settings
            MyGlobalSyncSettings.SourceDirectory = txtSourceDirectory.Text
            MyGlobalSyncSettings.ffmpegPath = txt_ffmpegPath.Text

            ' Save settings to file
            Dim MyResult As ReturnObject = SaveSyncSettings()

            If MyResult.Success Then
                MyLog.Write("Syncer settings updated.", Information)
                Me.DialogResult = True
                Me.Close()
            Else
                Throw New Exception(MyResult.ErrorMessage)
            End If
        Catch ex As Exception
            MyLog.Write("Could not update sync settings. Error: " & ex.Message, Warning)
            System.Windows.MessageBox.Show(ex.Message, "Save failed!", MessageBoxButton.OK, MessageBoxImage.Error)
            EnableDisableControls(True)
        End Try

    End Sub

    Private Sub btnCancel_Click(sender As Object, e As RoutedEventArgs)
        EnableDisableControls(False)
        Me.DialogResult = False
        Me.Close()
    End Sub

    Private Sub EnableDisableControls(Enable As Boolean)
        btnSave.IsEnabled = Enable
        btnCancel.IsEnabled = Enable
        txtSourceDirectory.IsEnabled = Enable
        txt_ffmpegPath.IsEnabled = Enable
        spinThreads.IsEnabled = Enable
    End Sub
#End Region

#Region " Window Closing "
    Private Sub MainWindow_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

    End Sub
#End Region

End Class
