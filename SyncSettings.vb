Public Class SyncSettings

    Property SyncIsEnabled As Boolean
    Property SourceDirectory As String
    Property SyncDirectory As String
    Dim WatcherCodecFilter As Codec()
    Dim WatcherTags As Codec.Tag()
    Property TranscodeLosslessFiles As Boolean?
    Property Encoder As Codec
    Property MaxThreads As Int32
    Property ffmpegPath As String

#Region " New "
    Public Sub New(MySyncIsEnabled As Boolean, MySourceDirectory As String, MySyncDirectory As String, MyWatcherCodecs As List(Of Codec),
                   MyWatcherTags As List(Of Codec.Tag), MyTranscodeLosslessFiles As Boolean, MyEncoder As Codec, MyMaxThreads As Int32, My_ffmpegPath As String)

        SyncIsEnabled = MySyncIsEnabled
        SourceDirectory = MySourceDirectory
        SyncDirectory = MySyncDirectory
        WatcherCodecFilter = MyWatcherCodecs.ToArray
        WatcherTags = MyWatcherTags.ToArray
        TranscodeLosslessFiles = MyTranscodeLosslessFiles
        Encoder = MyEncoder
        MaxThreads = MyMaxThreads
        ffmpegPath = My_ffmpegPath

    End Sub

    Public Sub New(NewSyncSettings As SyncSettings)

        If Not NewSyncSettings Is Nothing Then
            SyncIsEnabled = NewSyncSettings.SyncIsEnabled
            SourceDirectory = NewSyncSettings.SourceDirectory
            SyncDirectory = NewSyncSettings.SyncDirectory
            WatcherCodecFilter = NewSyncSettings.GetWatcherCodecs
            WatcherTags = NewSyncSettings.GetWatcherTags
            TranscodeLosslessFiles = NewSyncSettings.TranscodeLosslessFiles
            Encoder = NewSyncSettings.Encoder
            MaxThreads = NewSyncSettings.MaxThreads
            ffmpegPath = NewSyncSettings.ffmpegPath
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
End Class
