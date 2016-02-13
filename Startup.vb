#Region " Namespaces "
Imports Music_Folder_Syncer.Logger.DebugLogLevel
Imports Music_Folder_Syncer.Toolkit
Imports System.IO
Imports System.Environment
#End Region

'======================================================================
'=================== FEATURES TO ADD / BUGS TO FIX ====================
' - Add album art copying (parse images based on priority, like Foobar)
'======================================================================

Module Startup

    Public Const ApplicationName As String = "Music Folder Syncer"
    Public MyLog As Logger
    Public MyLogFilePath As String = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, ApplicationName & ".log")
    Public Const DebugLevel As Logger.DebugLogLevel = Information
    Public MySyncSettings As SyncSettings
    Public DefaultSyncSettings As SyncSettings
    Public Codecs As List(Of Codec)
    Public Const MaxFileID As Int32 = 99999

    Sub Main()

        MyLog = New Logger(MyLogFilePath)

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
            DefaultSyncSettings = DirectCast(DefaultSettings.MyObject, SyncSettings)
        Else
            MessageBox.Show(DefaultSettings.ErrorMessage, "Default Sync Settings Error", MessageBoxButton.OK, MessageBoxImage.Error)
            Exit Sub
        End If

        ' Read SyncSettings.xml file to import current sync settings (if there is one)
        Dim Settings As ReturnObject = XML.ReadSyncSettings(Codecs, DefaultSyncSettings)
        If Settings.Success Then
            If Settings.MyObject Is Nothing Then
                MyLog.Write("Settings file not found; launching new sync window.", Information)
                MessageBox.Show("No existing sync setup was found. Please create one now.",
                                "No Sync Settings Found", MessageBoxButton.OK, MessageBoxImage.Information)
                System.Windows.Forms.Application.Run(New TrayApp(True))
            Else
                MySyncSettings = DirectCast(Settings.MyObject, SyncSettings)
                Forms.Application.Run(New TrayApp(False))
            End If
        Else
            MessageBox.Show(Settings.ErrorMessage, "Sync Settings Error", MessageBoxButton.OK, MessageBoxImage.Error)
            Exit Sub
        End If

    End Sub

#Region " CLASS - FileParser "
    Class FileParser

#Region " Transfer File To Sync Folder "
        Private ProcessID As Int32
        ReadOnly Property FilePath As String

        Public Sub New(ByVal NewProcessID As Int32, ByVal NewFilePath As String)
            ProcessID = NewProcessID
            FilePath = NewFilePath
        End Sub

        Public Function TransferToSyncFolder(ByRef CodecsToCheck As Codec()) As ReturnObject

            Dim FileCodec As Codec = CheckFileCodec(CodecsToCheck)
            Dim MyReturnObject As ReturnObject

            If Not FileCodec Is Nothing Then
                Try
                    If CheckFileForSync(FileCodec) Then
                        Dim SyncFilePath As String = MySyncSettings.SyncDirectory & FilePath.Substring(MySyncSettings.SourceDirectory.Length)

                        If MySyncSettings.TranscodeLosslessFiles AndAlso FileCodec.Type = Codec.CodecType.Lossless Then 'Need to transcode file
                            MyLog.Write(ProcessID, "...transcoding file to " & MySyncSettings.Encoder.Name & "...", Debug)
                            TranscodeFile(SyncFilePath)

                            SyncFilePath = Path.Combine(Path.GetDirectoryName(SyncFilePath), Path.GetFileNameWithoutExtension(SyncFilePath)) &
                                MySyncSettings.Encoder.FileExtensions(0)
                        Else
                            Directory.CreateDirectory(Path.GetDirectoryName(SyncFilePath))
                            File.Copy(FilePath, SyncFilePath, True)
                        End If

                        Dim NewFile As New FileInfo(SyncFilePath)
                        'Interlocked.Add(SyncFolderSize, NewFile.Length)
                        MyLog.Write(ProcessID, "...successfully added file to sync folder...", Debug)
                        MyReturnObject = New ReturnObject(True, "", NewFile.Length)
                    Else
                        MyReturnObject = New ReturnObject(True, "", 0)
                    End If

                    MyLog.Write(ProcessID, "File processed: """ & FilePath.Substring(MySyncSettings.SourceDirectory.Length) & """", Information)
                Catch ex As Exception
                    MyLog.Write(ProcessID, "Processing failed: """ & FilePath.Substring(MySyncSettings.SourceDirectory.Length) & """. Exception: " & ex.Message, Warning)
                    MyReturnObject = New ReturnObject(False, ex.Message, 0)
                End Try
            Else
                MyLog.Write(ProcessID, "Ignoring file: """ & FilePath.Substring(MySyncSettings.SourceDirectory.Length) & """", Information)
                MyReturnObject = New ReturnObject(True, "", 0)
            End If

            Return MyReturnObject

        End Function

        Public Sub TranscodeFile(FileTo As String)

            Dim FileFrom As String = FilePath
            Dim OutputFilePath As String = ""

            Try
                Dim OutputDirectory As String = Path.GetDirectoryName(FileTo)
                OutputFilePath = Path.Combine(OutputDirectory, Path.GetFileNameWithoutExtension(FileTo)) & MySyncSettings.Encoder.FileExtensions(0)
                Directory.CreateDirectory(OutputDirectory)
            Catch ex As Exception
                MyLog.Write(ProcessID, "...transcode failed [1]. Exception: " & ex.Message & NewLine & NewLine & ex.InnerException.ToString, Warning)
            End Try

            Try
                Dim ffmpeg As New ProcessStartInfo(MySyncSettings.ffmpegPath)
                ffmpeg.CreateNoWindow = True
                ffmpeg.UseShellExecute = False

                ffmpeg.Arguments = "-i """ & FileFrom & """ -vn -c:a " & MySyncSettings.Encoder.Profiles(0).Argument & " -hide_banner """ & OutputFilePath & """"
                'libvorbis -aq: 4 = 128 kbps, 5 = 160 kbps, 6 = 192 kbps, 7 = 224 kbps, 8 = 256 kbps

                MyLog.Write(ProcessID, "...ffmpeg arguments: """ & ffmpeg.Arguments & """...", Debug)

                Dim ffmpegProcess As Process = Process.Start(ffmpeg)
                ffmpegProcess.WaitForExit()

                If ffmpegProcess.ExitCode <> 0 Then
                    Throw New Exception("ffmpeg exited with an error! (Code: " & ffmpegProcess.ExitCode & ")")
                End If

                MyLog.Write(ProcessID, "...transcode complete...", Debug)
            Catch ex As Exception
                MyLog.Write(ProcessID, "...transcode failed [2]. Exception: " & ex.Message & NewLine & NewLine & ex.InnerException.ToString, Warning)
            End Try

        End Sub
#End Region

#Region " File Checks "
        Public Function CheckFileForSync(ByVal FileCodec As Codec) As Boolean

            Try
                If CheckFileTags(FileCodec) Then
                    MyLog.Write(ProcessID, "...file has correct tags, now syncing...", Debug)
                    Return True
                Else
                    MyLog.Write(ProcessID, "...file does not have correct tags, ignoring...", Debug)
                    Return False
                End If
            Catch ex As Exception
                MyLog.Write(ProcessID, "...error whilst attempting to parse file. Exception: " & ex.Message, Warning)
                Return False
            End Try

        End Function

        Public Function CheckFileCodec(CodecsToCheck As Codec()) As Codec

            Dim FileExtension As String = Path.GetExtension(FilePath)

            For Each Codec As Codec In CodecsToCheck
                For Each CodecExtension As String In Codec.FileExtensions
                    If FileExtension = CodecExtension Then
                        MyLog.Write(ProcessID, "...file type recognised, now checking tags...", Debug)
                        Return Codec
                    End If
                Next
            Next

            MyLog.Write(ProcessID, "...file type not recognised, ignoring...", Debug)
            Return Nothing

        End Function

        Private Function CheckFileTags(MyCodec As Codec) As Boolean

            Dim TagsObject As ReturnObject = MyCodec.MatchTag(FilePath, MySyncSettings.GetWatcherTags)

            If TagsObject.Success Then
                Return CType(TagsObject.MyObject, Boolean)
            Else
                MyLog.Write(ProcessID, "...could not obtain file tags. File: """ & FilePath & """, Codec: " & MyCodec.Name & ", Exception: " & TagsObject.ErrorMessage, Warning)
                Return False
            End If

        End Function
#End Region

    End Class
#End Region

End Module
