Public Class GlobalSyncSettings

    Property SyncIsEnabled As Boolean
    Property SourceDirectory As String
    Property ffmpegPath As String
    Property LogLevel As Logger.LogLevel
    Dim SyncSettings As SyncSettings()

#Region " New "
    Public Sub New(MySyncIsEnabled As Boolean, MySourceDirectory As String, My_ffmpegPath As String, MySyncSettingsList As List(Of SyncSettings), MyLogLevel As String)

        SyncIsEnabled = MySyncIsEnabled
        SourceDirectory = MySourceDirectory
        ffmpegPath = My_ffmpegPath
        If Not MySyncSettingsList Is Nothing Then SyncSettings = MySyncSettingsList.ToArray()
        LogLevel = Logger.ConvertStringToLogLevel(MyLogLevel)

    End Sub

    Public Sub New(NewSyncSettings As GlobalSyncSettings)

        If Not NewSyncSettings Is Nothing Then
            SyncIsEnabled = NewSyncSettings.SyncIsEnabled
            SourceDirectory = NewSyncSettings.SourceDirectory
            ffmpegPath = NewSyncSettings.ffmpegPath
            SyncSettings = NewSyncSettings.GetSyncSettings()
        End If

    End Sub
#End Region

    Public Function GetSyncSettings() As SyncSettings()
        Return SyncSettings
    End Function

    Public Sub SetSyncSettings(SyncSettingsList As List(Of SyncSettings))
        If Not SyncSettingsList Is Nothing Then SyncSettings = SyncSettingsList.ToArray()
    End Sub

    Public Sub SetLogLevel(NewLogLevel As String)
        LogLevel = Logger.ConvertStringToLogLevel(NewLogLevel)
    End Sub

    Public Function GetLogLevel() As String
        Return Logger.ConvertLogLevelToString(LogLevel)
    End Function

End Class


Public Class SyncSettings

    Property SyncDirectory As String
    Dim WatcherCodecFilter As Codec()
    Dim WatcherTags As Codec.Tag()
    Property TranscodeSetting As TranscodeMode
    Property ReplayGainSetting As ReplayGainMode
    Property Encoder As Codec
    Property MaxThreads As Int32

    Enum ReplayGainMode
        None
        Track
        Album
    End Enum

    Enum TranscodeMode
        None
        LosslessOnly
        All
    End Enum

#Region " New "
    Public Sub New(MySyncDirectory As String, MyWatcherCodecs As List(Of Codec), MyWatcherTags As List(Of Codec.Tag),
                   MyTranscodeLosslessFiles As TranscodeMode, MyEncoder As Codec, MyMaxThreads As Int32, MyReplayGain As ReplayGainMode)

        SyncDirectory = MySyncDirectory
        WatcherCodecFilter = MyWatcherCodecs.ToArray
        WatcherTags = MyWatcherTags.ToArray
        TranscodeSetting = MyTranscodeLosslessFiles
        Encoder = MyEncoder
        MaxThreads = MyMaxThreads
        ReplayGainSetting = MyReplayGain

    End Sub

    Public Sub New(NewSyncSettings As SyncSettings)

        If Not NewSyncSettings Is Nothing Then
            SyncDirectory = NewSyncSettings.SyncDirectory
            WatcherCodecFilter = NewSyncSettings.GetWatcherCodecs
            WatcherTags = NewSyncSettings.GetWatcherTags
            TranscodeSetting = NewSyncSettings.TranscodeSetting
            Encoder = NewSyncSettings.Encoder
            MaxThreads = NewSyncSettings.MaxThreads
            ReplayGainSetting = NewSyncSettings.ReplayGainSetting
        End If

    End Sub
#End Region

    Public Function GetFileExtensions() As String() 'List(Of String)

        Dim FileExtensions As New List(Of String)

        For Each MyCodec As Codec In WatcherCodecFilter
            For Each MyExtension As String In MyCodec.GetFileExtensions()
                FileExtensions.Add(MyExtension)
            Next
        Next

        Return FileExtensions.ToArray

    End Function

    Public Function GetWatcherCodecs() As Codec()
        Return WatcherCodecFilter
    End Function

    Public Sub SetWatcherCodecs(CodecList As List(Of Codec))
        If Not CodecList Is Nothing Then WatcherCodecFilter = CodecList.ToArray
    End Sub

    Public Function GetWatcherTags() As Codec.Tag()
        Return WatcherTags
    End Function

    Public Sub SetWatcherTags(TagList As List(Of Codec.Tag))
        If Not TagList Is Nothing Then WatcherTags = TagList.ToArray
    End Sub

    Public Function GetReplayGainSetting() As String
        Return GetReplayGainSetting(ReplayGainSetting)
    End Function

    Public Function GetTranscodeSetting() As String
        Return GetTranscodeSetting(TranscodeSetting)
    End Function

    Public Sub SetReplayGainSetting(ReplayGainSettingString As String)
        Select Case ReplayGainSettingString
            Case Is = "Album"
                ReplayGainSetting = ReplayGainMode.Album
            Case Is = "Track"
                ReplayGainSetting = ReplayGainMode.Track
            Case Else
                ReplayGainSetting = ReplayGainMode.None
        End Select
    End Sub

    Public Sub SetTranscodeSetting(TranscodeSettingString As String)
        Select Case TranscodeSettingString
            Case Is = "Lossless Only"
                TranscodeSetting = TranscodeMode.LosslessOnly
            Case Is = "All"
                TranscodeSetting = TranscodeMode.All
            Case Else
                TranscodeSetting = TranscodeMode.None
        End Select
    End Sub

    Public Shared Function GetReplayGainSetting(ReplayGainSetting As ReplayGainMode) As String
        Select Case ReplayGainSetting
            Case ReplayGainMode.Album
                Return "Album"
            Case ReplayGainMode.Track
                Return "Track"
            Case Else
                Return "None"
        End Select
    End Function

    Public Shared Function GetTranscodeSetting(TranscodeSetting As TranscodeMode) As String
        Select Case TranscodeSetting
            Case TranscodeMode.LosslessOnly
                Return "Lossless Only"
            Case TranscodeMode.All
                Return "All"
            Case Else
                Return "None"
        End Select
    End Function

    'Public Shared Function SetTranscodeSetting(ReplayGainSettingString As String) As ReplayGainMode
    '    Select Case ReplayGainSettingString
    '        Case Is = "Album"
    '            Return ReplayGainMode.Album
    '        Case Is = "Track"
    '            Return ReplayGainMode.Track
    '        Case Else
    '            Return ReplayGainMode.None
    '    End Select
    'End Function

End Class
