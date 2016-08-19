#Region " Namespaces "
Imports MusicFolderSyncer.Logger.LogLevel
Imports MusicFolderSyncer.Toolkit
Imports System.IO
Imports Microsoft.WindowsAPICodePack.Dialogs
#End Region

'======================================================================
'=================== FEATURES TO ADD / BUGS TO FIX ====================
' - Add album art copying (parse images based on priority, like Foobar)
'======================================================================

Module Startup

    Public EnglishGB As System.Globalization.CultureInfo = System.Globalization.CultureInfo.CreateSpecificCulture("en-GB")
    Private Const LogLevel As Logger.LogLevel = Information
    Public MyLogFilePath As String = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, ApplicationName & ".log")

    Public Const ApplicationName As String = "Music Folder Syncer"
    Public MyLog As Logger
    Public UserGlobalSyncSettings As GlobalSyncSettings = Nothing
    Public DefaultGlobalSyncSettings As GlobalSyncSettings = Nothing
    Public Codecs As List(Of Codec)
    Public Const MaxFileID As Int32 = 99999

#Region " Sub Main "
    Sub Main()

        MyLog = New Logger(MyLogFilePath, LogLevel)

        MyLog.Write("===============================================================")
        MyLog.Write("  PROGRAM LAUNCHED")
        MyLog.Write("===============================================================")

        ' Read Codecs.xml file to import list of recognised codecs
        Codecs = XML.ReadCodecs()
        If Codecs Is Nothing Then
            MessageBox.Show("Could not read from Codecs.xml! Ensure the file is present and in the correct format.",
                            "Codecs Error", MessageBoxButton.OK, MessageBoxImage.Error)
            Exit Sub
        End If

        ' Read DefaultSettings.xml file to import default sync settings
        Dim DefaultSettings As ReturnObject = XML.ReadDefaultSettings(Codecs)
        If DefaultSettings.Success AndAlso Not DefaultSettings.MyObject Is Nothing Then
            DefaultGlobalSyncSettings = DirectCast(DefaultSettings.MyObject, GlobalSyncSettings)
        Else
            MessageBox.Show(DefaultSettings.ErrorMessage, "Default Sync Settings Error", MessageBoxButton.OK, MessageBoxImage.Error)
            Exit Sub
        End If

        MyLog.DebugLevel = DefaultGlobalSyncSettings.LogLevel

        ' Read SyncSettings.xml file to import current sync settings (if there is one)
        Dim Settings As ReturnObject = XML.ReadSyncSettings(Codecs, DefaultGlobalSyncSettings)
        If Settings.Success Then
            If Settings.MyObject Is Nothing Then
                MyLog.Write("Settings file not found; launching new sync window.", Information)
                MessageBox.Show("No existing sync setup was found. Please create one now.",
                                "No Sync Settings Found", MessageBoxButton.OK, MessageBoxImage.Information)
                Forms.Application.Run(New TrayApp(True))
            Else
                UserGlobalSyncSettings = DirectCast(Settings.MyObject, GlobalSyncSettings)
                MyLog.DebugLevel = UserGlobalSyncSettings.LogLevel
                MyLog.Write("Changing log level to " & Logger.ConvertLogLevelToString(UserGlobalSyncSettings.LogLevel), Logger.LogLevel.Debug)
                Forms.Application.Run(New TrayApp(False))
            End If
        Else
            MessageBox.Show(Settings.ErrorMessage, "Sync Settings Error", MessageBoxButton.OK, MessageBoxImage.Error)
            Exit Sub
        End If

    End Sub
#End Region

    Public Function CreateDirectoryBrowser(StartingDirectory As String) As ReturnObject

        Using SelectDirectoryDialog = New CommonOpenFileDialog()
            SelectDirectoryDialog.Title = "Select Sync Directory"
            SelectDirectoryDialog.IsFolderPicker = True
            SelectDirectoryDialog.AddToMostRecentlyUsedList = False
            SelectDirectoryDialog.AllowNonFileSystemItems = True
            SelectDirectoryDialog.DefaultDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            If Directory.Exists(StartingDirectory) Then
                SelectDirectoryDialog.InitialDirectory = StartingDirectory
            Else
                SelectDirectoryDialog.InitialDirectory = SelectDirectoryDialog.DefaultDirectory
            End If
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
        End Using

    End Function

    Public Function CreateFileBrowser(DefaultPath As String) As ReturnObject

        Using SelectDirectoryDialog = New CommonOpenFileDialog()
            SelectDirectoryDialog.Title = "Select ffmpeg.exe Path"
            SelectDirectoryDialog.IsFolderPicker = False
            SelectDirectoryDialog.AddToMostRecentlyUsedList = False
            SelectDirectoryDialog.AllowNonFileSystemItems = True

            SelectDirectoryDialog.DefaultDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
            If File.Exists(DefaultPath) Then
                SelectDirectoryDialog.InitialDirectory = Path.GetDirectoryName(DefaultPath)
            Else
                SelectDirectoryDialog.InitialDirectory = SelectDirectoryDialog.DefaultDirectory
            End If

            SelectDirectoryDialog.EnsureFileExists = True
            SelectDirectoryDialog.EnsurePathExists = True
            SelectDirectoryDialog.EnsureReadOnly = False
            SelectDirectoryDialog.EnsureValidNames = True
            SelectDirectoryDialog.Multiselect = False
            SelectDirectoryDialog.ShowPlacesList = True
            SelectDirectoryDialog.Filters.Add(New CommonFileDialogFilter("ffmpeg Executable", "*.exe"))

            If SelectDirectoryDialog.ShowDialog() = CommonFileDialogResult.Ok Then
                Return New ReturnObject(True, "", SelectDirectoryDialog.FileName)
            Else
                Return New ReturnObject(False, "")
            End If
        End Using

    End Function

End Module
