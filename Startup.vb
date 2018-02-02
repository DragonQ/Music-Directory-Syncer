#Region " Namespaces "
Imports MusicFolderSyncer.Logger.LogLevel
Imports MusicFolderSyncer.Toolkit
Imports System.IO
Imports System.Security.AccessControl
Imports System.Security.Principal
#End Region

Module Startup

    Public EnglishGB As System.Globalization.CultureInfo = System.Globalization.CultureInfo.CreateSpecificCulture("en-GB")
    Private Const LogLevel As Logger.LogLevel = Information
    Public MyLogFilePath As String = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, ApplicationName & ".log")
    Public MyVersion As String = ""
    Public MyDirectoryPermissions As DirectorySecurity

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

        'Create global directory access permissions
        Dim CurrentUser As WindowsIdentity = WindowsIdentity.GetCurrent()
        Dim SystemUser = New SecurityIdentifier(WellKnownSidType.LocalSystemSid, Nothing)
        Dim OtherUsers = New SecurityIdentifier(WellKnownSidType.BuiltinUsersSid, Nothing)
        MyDirectoryPermissions = New DirectorySecurity
        MyDirectoryPermissions.AddAccessRule(New FileSystemAccessRule(CurrentUser.User,
                                                                      FileSystemRights.FullControl,
                                                                      InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit,
                                                                      PropagationFlags.None,
                                                                      AccessControlType.Allow))
        MyDirectoryPermissions.AddAccessRule(New FileSystemAccessRule(SystemUser,
                                                                      FileSystemRights.FullControl,
                                                                      InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit,
                                                                      PropagationFlags.None,
                                                                      AccessControlType.Allow))
        MyDirectoryPermissions.AddAccessRule(New FileSystemAccessRule(OtherUsers,
                                                                      FileSystemRights.ReadAndExecute,
                                                                      InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit,
                                                                      PropagationFlags.None,
                                                                      AccessControlType.Allow))

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

End Module