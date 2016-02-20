#Region " Namespaces "
Imports MusicFolderSyncer.Logger.DebugLogLevel
Imports MusicFolderSyncer.Toolkit
Imports System.IO
Imports System.Environment
#End Region

Class FileParser

    Private ProcessID As Int32
    ReadOnly Property FilePath As String

    Public Sub New(ByVal NewProcessID As Int32, ByVal NewFilePath As String)
        ProcessID = NewProcessID
        FilePath = NewFilePath
    End Sub

#Region " Transfer File To Sync Folder "
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
                            MySyncSettings.Encoder.GetFileExtensions(0)
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
            OutputFilePath = Path.Combine(OutputDirectory, Path.GetFileNameWithoutExtension(FileTo)) & MySyncSettings.Encoder.GetFileExtensions(0)
            Directory.CreateDirectory(OutputDirectory)
        Catch ex As Exception
            MyLog.Write(ProcessID, "...transcode failed [1]. Exception: " & ex.Message & NewLine & NewLine & ex.InnerException.ToString, Warning)
        End Try

        Try
            Dim ffmpeg As New ProcessStartInfo(MySyncSettings.ffmpegPath)
            ffmpeg.CreateNoWindow = True
            ffmpeg.UseShellExecute = False

            ffmpeg.Arguments = "-i """ & FileFrom & """ -vn -c:a " & MySyncSettings.Encoder.GetProfiles(0).Argument & " -hide_banner """ & OutputFilePath & """"
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
            For Each CodecExtension As String In Codec.GetFileExtensions()
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