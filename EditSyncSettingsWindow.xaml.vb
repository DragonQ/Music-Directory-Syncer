Imports Music_Folder_Syncer.Logger.DebugLogLevel


Public Class EditSyncSettingsWindow

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        txtSourceDirectory.Text = MySyncSettings.SourceDirectory
        txtSyncDirectory.Text = MySyncSettings.SyncDirectory
        spinThreads.Maximum = System.Environment.ProcessorCount
        spinThreads.Value = MySyncSettings.MaxThreads

    End Sub

#Region " Window Controls "
    Private Sub btnSave_Click(sender As Object, e As RoutedEventArgs)

        EnableDisableControls(False)
        Dim MyResult As ReturnObject = SaveSyncSettings()

        If MyResult.Success Then
            MyLog.Write("Syncer settings updated.", Information)
            Me.DialogResult = True
            Me.Close()
        Else
            MyLog.Write("Could not update sync settings. Error: " & MyResult.ErrorMessage, Warning)
            MessageBox.Show("Save failed!", MyResult.ErrorMessage, MessageBoxButton.OK, MessageBoxImage.Error)
            EnableDisableControls(True)
        End If

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
