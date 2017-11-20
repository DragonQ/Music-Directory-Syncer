#Region " Namespaces "
Imports MusicFolderSyncer.Logger.LogLevel
Imports MusicFolderSyncer.Toolkit
Imports System.IO
Imports Microsoft.WindowsAPICodePack.Dialogs
#End Region

Module Startup

    Public EnglishGB As System.Globalization.CultureInfo = System.Globalization.CultureInfo.CreateSpecificCulture("en-GB")
    Private Const LogLevel As Logger.LogLevel = Information
    Public MyLogFilePath As String = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, ApplicationName & ".log")
    Public MyVersion As String = ""

    Public Const ApplicationName As String = "Music Folder Syncer"
    Public MyLog As Logger
    Public UserGlobalSyncSettings As GlobalSyncSettings = Nothing
    Public DefaultGlobalSyncSettings As GlobalSyncSettings = Nothing
    Public Codecs As List(Of Codec)
    Public Const MaxFileID As Int32 = 99999

#Region " Sub Main "
    Sub Main()

        ' Create public logger for entire program to use
        MyLog = New Logger(MyLogFilePath, LogLevel)

        ' Grab version number and write to log
        MyVersion = ThisAssembly.Git.BaseTag
        If ThisAssembly.Git.IsDirty Then
            MyVersion &= "." & ThisAssembly.Git.Commits & "-" & ThisAssembly.Git.Commit
        End If
        MyLog.Write("===============================================================")
        MyLog.Write("  " & ApplicationName & " " & MyVersion & " - STARTUP")
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
            With SelectDirectoryDialog
                .Title = "Select Sync Directory"
                .IsFolderPicker = True
                .AddToMostRecentlyUsedList = False
                .AllowNonFileSystemItems = True
                .DefaultDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                If Directory.Exists(StartingDirectory) Then
                    .InitialDirectory = StartingDirectory
                Else
                    .InitialDirectory = .DefaultDirectory
                End If
                .EnsureFileExists = False
                .EnsurePathExists = True
                .EnsureReadOnly = False
                .EnsureValidNames = True
                .Multiselect = False
                .ShowPlacesList = True
            End With

            If SelectDirectoryDialog.ShowDialog() = CommonFileDialogResult.Ok Then
                Return New ReturnObject(True, "", SelectDirectoryDialog.FileName)
            Else
                Return New ReturnObject(False, "")
            End If
        End Using

    End Function

    Public Function CreateFileBrowser(DefaultPath As String) As ReturnObject

        Using SelectFileDialog = New CommonOpenFileDialog()
            With SelectFileDialog
                .Title = "Select ffmpeg.exe Path"
                .IsFolderPicker = False
                .AddToMostRecentlyUsedList = False
                .AllowNonFileSystemItems = True

                .DefaultDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
                If File.Exists(DefaultPath) Then
                    .InitialDirectory = Path.GetDirectoryName(DefaultPath)
                Else
                    .InitialDirectory = .DefaultDirectory
                End If

                .EnsureFileExists = True
                .EnsurePathExists = True
                .EnsureReadOnly = False
                .EnsureValidNames = True
                .Multiselect = False
                .ShowPlacesList = True
                .Filters.Add(New CommonFileDialogFilter("ffmpeg Executable", "*.exe"))
            End With

            If SelectFileDialog.ShowDialog() = CommonFileDialogResult.Ok Then
                Return New ReturnObject(True, "", SelectFileDialog.FileName)
            Else
                Return New ReturnObject(False, "")
            End If
        End Using

    End Function

End Module
