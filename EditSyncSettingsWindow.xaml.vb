#Region " Namespaces "
Imports Music_Folder_Syncer.Logger.DebugLogLevel
Imports Music_Folder_Syncer.Toolkit
Imports Microsoft.WindowsAPICodePack.Dialogs
#End Region

Public Class EditSyncSettingsWindow

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

    End Sub

    Private Sub EditSyncSettingsWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

        ' Set run-time properties of window objects
        txtSourceDirectory.Text = MySyncSettings.SourceDirectory
        txtSyncDirectory.Text = MySyncSettings.SyncDirectory
        spinThreads.Maximum = Environment.ProcessorCount
        spinThreads.Value = MySyncSettings.MaxThreads

    End Sub

    Private Function CreateDirectoryBrowser() As ReturnObject

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
            Return New ReturnObject(True, "", SelectDirectoryDialog.FileName)
        Else
            Return New ReturnObject(False, "")
        End If

    End Function

#Region " Window Controls "
    Private Sub btnBrowseSourceDirectory_Click(sender As Object, e As RoutedEventArgs)

        Dim Browser As ReturnObject = CreateDirectoryBrowser()

        If Browser.Success Then
            txtSourceDirectory.Text = CStr(Browser.MyObject)
        End If

    End Sub

    Private Sub btnBrowseSyncDirectory_Click(sender As Object, e As RoutedEventArgs)

        Dim Browser As ReturnObject = CreateDirectoryBrowser()

        If Browser.Success Then
            txtSyncDirectory.Text = CStr(Browser.MyObject)
        End If

    End Sub

    Private Sub btnSave_Click(sender As Object, e As RoutedEventArgs)

        EnableDisableControls(False)

        Try
            ' Check settings aren't nonsense
            If Not IO.Directory.Exists(txtSourceDirectory.Text) Then
                Throw New Exception("Specified source directory doesn't exist.")
            End If

            If Not IO.Directory.Exists(txtSyncDirectory.Text) Then
                If (IO.Path.IsPathRooted(txtSyncDirectory.Text) AndAlso IO.Path.IsPathRooted(txtSyncDirectory.Text)) Then
                    IO.Directory.CreateDirectory(txtSyncDirectory.Text)
                Else
                    Throw New Exception("Specified sync directory is not valid.")
                End If
            End If

            ' Apply settings
            MySyncSettings.SourceDirectory = txtSourceDirectory.Text
            MySyncSettings.SyncDirectory = txtSyncDirectory.Text
            MySyncSettings.MaxThreads = CInt(spinThreads.Value)

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
        txtSyncDirectory.IsEnabled = Enable
        spinThreads.IsEnabled = Enable
    End Sub
#End Region

#Region " Window Closing "
    Private Sub MainWindow_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

    End Sub
#End Region

End Class
