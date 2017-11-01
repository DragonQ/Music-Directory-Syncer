#Region " Namespaces "
Imports MusicFolderSyncer.Logger.LogLevel
Imports MusicFolderSyncer.Toolkit
Imports MusicFolderSyncer.SyncSettings
Imports MusicFolderSyncer.Codec.CodecType
Imports MusicFolderSyncer.SyncSettings.TranscodeMode
Imports System.IO
Imports System.Environment
Imports System.Threading
#End Region

Class FileParser

    Implements IDisposable

    Private ProcessID As Int32
    Private Shared FileTimeout As Int32 = 60
    ReadOnly Property FilePath As String
    Private FileLock As Boolean
    Private MyGlobalSyncSettings As GlobalSyncSettings
    Private SyncSettings As SyncSettings()
    Private SourceFileStream As FileStream = Nothing

#Region " New "
    Public Sub New(ByRef NewGlobalSyncSettings As GlobalSyncSettings, ByVal NewProcessID As Int32, ByVal NewFilePath As String, Optional NewSyncSettings As SyncSettings() = Nothing)

        ProcessID = NewProcessID
        FilePath = NewFilePath
        MyGlobalSyncSettings = NewGlobalSyncSettings
        If NewSyncSettings Is Nothing Then
            SyncSettings = MyGlobalSyncSettings.GetSyncSettings()
        Else
            SyncSettings = NewSyncSettings
        End If

        If File.Exists(FilePath) Then
            SourceFileStream = WaitForFile(FilePath, FileMode.Open, FileAccess.Read, FileShare.Read, FileTimeout)
            If SourceFileStream Is Nothing Then
                MyLog.Write(ProcessID, "Could not get file system lock on source file: """ & FilePath.Substring(MyGlobalSyncSettings.SourceDirectory.Length) & """.", Warning)
                FileLock = False
            Else
                FileLock = True
            End If
        Else
            MyLog.Write(ProcessID, "Source file doesn't exist: """ & FilePath.Substring(MyGlobalSyncSettings.SourceDirectory.Length) & """.", Debug)
            FileLock = False
        End If

    End Sub
#End Region

    Public Overridable Sub Dispose() Implements IDisposable.Dispose
        If Not SourceFileStream Is Nothing Then
            SourceFileStream.Close()
        End If
    End Sub

#Region " Transfer File To Sync Folder "
    Public Function DeleteInSyncFolder() As ReturnObject

        Dim MyReturnObject As ReturnObject

        Try
            For Each SyncSetting In SyncSettings
                Dim FileCodec As Codec = CheckFileCodec(SyncSetting.GetWatcherCodecs())
                If Not FileCodec Is Nothing Then
                    'File was meant to be synced, which means we now need to delete the synced version
                    Dim SyncFilePath As String = SyncSetting.SyncDirectory & FilePath.Substring(MyGlobalSyncSettings.SourceDirectory.Length)

                    If SyncSetting.TranscodeSetting = All OrElse (SyncSetting.TranscodeSetting = LosslessOnly AndAlso FileCodec.CompressionType = Lossless) Then 'Need to replace extension with .ogg
                        Dim TranscodedFilePath As String = Path.Combine(Path.GetDirectoryName(SyncFilePath), Path.GetFileNameWithoutExtension(SyncFilePath)) &
                                                        SyncSetting.Encoder.GetFileExtensions(0)
                        SyncFilePath = TranscodedFilePath
                    End If

                    'Delete file if it exists in sync folder
                    If File.Exists(SyncFilePath) Then
                        File.Delete(SyncFilePath)
                        MyLog.Write(ProcessID, "...file in sync folder deleted: """ & SyncFilePath.Substring(SyncSetting.SyncDirectory.Length) & """.", Information)
                    Else
                        MyLog.Write(ProcessID, "...file doesn't exist in sync folder: """ & SyncFilePath.Substring(SyncSetting.SyncDirectory.Length) & """.", Information)
                    End If

                    'Delete folder if there are no more files in it
                    Dim DirName = Path.GetDirectoryName(SyncFilePath)
                    If Directory.Exists(DirName) AndAlso Directory.GetFiles(DirName, "*", SearchOption.AllDirectories).Length = 0 Then
                        Directory.Delete(DirName)
                        MyLog.Write(ProcessID, "...parent directory is now empty so was deleted: """ & DirName.Substring(SyncSetting.SyncDirectory.Length) & """.", Information)
                    End If
                Else
                    Throw New Exception("File was being watched but could not determine its codec.")
                End If
            Next

            MyReturnObject = New ReturnObject(True, "", 0)
            MyLog.Write(ProcessID, "Sync file deleted: """ & FilePath.Substring(MyGlobalSyncSettings.SourceDirectory.Length) & """", Information)
        Catch ex As Exception
            MyLog.Write(ProcessID, "Sync file deletion failed: """ & FilePath.Substring(MyGlobalSyncSettings.SourceDirectory.Length) & """. Exception: " & ex.Message, Warning)
            MyReturnObject = New ReturnObject(False, ex.Message, 0)
        End Try

        Return MyReturnObject

    End Function

    Public Function DeleteDirectoryInSyncFolder() As ReturnObject

        Dim MyReturnObject As ReturnObject

        Try
            For Each SyncSetting In SyncSettings
                'File was meant to be synced, which means we now need to delete the synced version
                Dim SyncFilePath As String = SyncSetting.SyncDirectory & FilePath.Substring(MyGlobalSyncSettings.SourceDirectory.Length)

                'Delete file if it exists in sync folder
                If Directory.Exists(SyncFilePath) Then
                    Directory.Delete(SyncFilePath, True)
                    MyLog.Write(ProcessID, "...directory in sync folder deleted: """ & SyncFilePath.Substring(SyncSetting.SyncDirectory.Length) & """.", Information)
                Else
                    MyLog.Write(ProcessID, "...directory doesn't exist in sync folder: """ & SyncFilePath.Substring(SyncSetting.SyncDirectory.Length) & """.", Information)
                End If
            Next

            MyReturnObject = New ReturnObject(True, "", 0)
            MyLog.Write(ProcessID, "Sync directory deleted: """ & FilePath.Substring(MyGlobalSyncSettings.SourceDirectory.Length) & """", Information)
        Catch ex As Exception
            MyLog.Write(ProcessID, "Sync directory deletion failed: """ & FilePath.Substring(MyGlobalSyncSettings.SourceDirectory.Length) & """. Exception: " & ex.Message, Warning)
            MyReturnObject = New ReturnObject(False, ex.Message, 0)
        End Try

        Return MyReturnObject

    End Function

    Public Function TransferToSyncFolder() As ReturnObject

        Dim MyReturnObject As ReturnObject

        Try
            If Not FileLock Then Throw New System.Exception("Could not get file system lock on source file.")
            Dim TranscodedFilePath As String = Nothing
            Dim NewFilesSize As Int64 = 0
            For Each SyncSetting In SyncSettings
                Dim FileCodec As Codec = CheckFileCodec(SyncSetting.GetWatcherCodecs())
                If Not FileCodec Is Nothing Then
                    If CheckFileForSync(FileCodec, SyncSetting) Then
                        Dim SyncFilePath As String = SyncSetting.SyncDirectory & FilePath.Substring(MyGlobalSyncSettings.SourceDirectory.Length)

                        If SyncSetting.TranscodeSetting = All OrElse (SyncSetting.TranscodeSetting = LosslessOnly AndAlso FileCodec.CompressionType = Lossless) Then 'Need to transcode file
                            If TranscodedFilePath Is Nothing Then 'File hasn't been transcoded for another sync, so transcode it
                                MyLog.Write(ProcessID, "...transcoding file to " & SyncSetting.Encoder.Name & "...", Debug)
                                Dim Result As ReturnObject = TranscodeFile(SyncFilePath, SyncSetting)
                                If Result.Success Then
                                    SyncFilePath = Path.Combine(Path.GetDirectoryName(SyncFilePath), Path.GetFileNameWithoutExtension(SyncFilePath)) &
                                                    SyncSetting.Encoder.GetFileExtensions(0)
                                    TranscodedFilePath = SyncFilePath
                                Else
                                    Throw New Exception(Result.ErrorMessage)
                                End If
                            Else
                                MyLog.Write(ProcessID, "...copying already transcoded file...", Debug)
                                SyncFilePath = Path.Combine(Path.GetDirectoryName(SyncFilePath), Path.GetFileNameWithoutExtension(SyncFilePath)) &
                                                SyncSetting.Encoder.GetFileExtensions(0)
                                Dim Result As ReturnObject = SafeCopy(TranscodedFilePath, SyncFilePath)
                                If Not Result.Success Then Throw New Exception(Result.ErrorMessage)
                            End If
                        Else
                            Directory.CreateDirectory(Path.GetDirectoryName(SyncFilePath))
                            Dim Result As ReturnObject = SafeCopy(SourceFileStream, SyncFilePath)
                            If Not Result.Success Then Throw New Exception(Result.ErrorMessage)
                        End If

                        Dim NewFile As New FileInfo(SyncFilePath)
                        NewFilesSize += NewFile.Length
                        MyLog.Write(ProcessID, "...successfully added file to sync folder...", Debug)
                    End If
                Else
                    MyLog.Write(ProcessID, "Ignoring file: """ & FilePath.Substring(MyGlobalSyncSettings.SourceDirectory.Length) & """", Information)
                End If
            Next

            MyReturnObject = New ReturnObject(True, "", NewFilesSize)
            MyLog.Write(ProcessID, "Sync file processed: """ & FilePath.Substring(MyGlobalSyncSettings.SourceDirectory.Length) & """", Information)
        Catch ex As Exception
            MyLog.Write(ProcessID, "Sync file processing failed: """ & FilePath.Substring(MyGlobalSyncSettings.SourceDirectory.Length) & """. Exception: " & ex.Message, Warning)
            MyReturnObject = New ReturnObject(False, ex.Message, 0)
        End Try

        Return MyReturnObject

    End Function

    Public Function RenameInSyncFolder(OldFilePath As String) As ReturnObject

        Dim MyReturnObject As ReturnObject

        Try
            If Not FileLock Then Throw New System.Exception("Could not get file system lock on source file.")
            Dim TranscodedFilePath As String = Nothing
            For Each SyncSetting In SyncSettings
                Dim FileCodec As Codec = CheckFileCodec(SyncSetting.GetWatcherCodecs())
                If Not FileCodec Is Nothing Then
                    If CheckFileForSync(FileCodec, SyncSetting) Then
                        Dim SyncFilePath As String = SyncSetting.SyncDirectory & FilePath.Substring(MyGlobalSyncSettings.SourceDirectory.Length)
                        Dim OldSyncFilePath As String = SyncSetting.SyncDirectory & OldFilePath.Substring(MyGlobalSyncSettings.SourceDirectory.Length)

                        If SyncSetting.TranscodeSetting = All OrElse (SyncSetting.TranscodeSetting = LosslessOnly AndAlso FileCodec.CompressionType = Lossless) Then 'Need to replace extension with .ogg
                            Dim TempString As String = Path.Combine(Path.GetDirectoryName(SyncFilePath), Path.GetFileNameWithoutExtension(SyncFilePath)) &
                                SyncSetting.Encoder.GetFileExtensions(0)
                            SyncFilePath = TempString
                            TempString = Path.Combine(Path.GetDirectoryName(OldSyncFilePath), Path.GetFileNameWithoutExtension(OldSyncFilePath)) &
                                SyncSetting.Encoder.GetFileExtensions(0)
                            OldSyncFilePath = TempString
                        End If

                        If File.Exists(OldSyncFilePath) Then
                            Dim DirName = Path.GetDirectoryName(SyncFilePath)
                            If Not Directory.Exists(DirName) Then Directory.CreateDirectory(DirName)
                            File.Move(OldSyncFilePath, SyncFilePath)
                            MyLog.Write(ProcessID, "...successfully renamed file in sync folder.", Information)

                            'Delete folder if there are no more files in it
                            Dim OldDirName = Path.GetDirectoryName(OldSyncFilePath)
                            If Directory.Exists(OldDirName) AndAlso Directory.GetFiles(OldDirName, "*", SearchOption.AllDirectories).Length = 0 Then
                                Directory.Delete(OldDirName)
                                MyLog.Write(ProcessID, "...old parent directory is now empty so was deleted: """ & OldDirName.Substring(SyncSetting.SyncDirectory.Length) & """.", Information)
                            End If
                        Else
                            MyLog.Write(ProcessID, "...old file doesn't exist in sync folder: """ & OldSyncFilePath & """, creating now...", Warning)

                            If SyncSetting.TranscodeSetting = All OrElse (SyncSetting.TranscodeSetting = LosslessOnly AndAlso FileCodec.CompressionType = Lossless) Then 'Need to transcode file
                                If TranscodedFilePath Is Nothing Then 'File hasn't been transcoded for another sync, so transcode it
                                    MyLog.Write(ProcessID, "...transcoding file to " & SyncSetting.Encoder.Name & "...", Debug)
                                    Dim Result As ReturnObject = TranscodeFile(SyncFilePath, SyncSetting)
                                    If Result.Success Then
                                        SyncFilePath = Path.Combine(Path.GetDirectoryName(SyncFilePath), Path.GetFileNameWithoutExtension(SyncFilePath)) &
                                                        SyncSetting.Encoder.GetFileExtensions(0)
                                        TranscodedFilePath = SyncFilePath
                                    Else
                                        Throw New Exception(Result.ErrorMessage)
                                    End If
                                Else 'File was transcoded before for a previous sync, so just copy that file
                                    MyLog.Write(ProcessID, "...copying already transcoded file...", Debug)
                                    Dim Result As ReturnObject = SafeCopy(TranscodedFilePath, SyncFilePath)
                                    If Not Result.Success Then Throw New Exception(Result.ErrorMessage)
                                    SyncFilePath = Path.Combine(Path.GetDirectoryName(SyncFilePath), Path.GetFileNameWithoutExtension(SyncFilePath)) &
                                                    SyncSetting.Encoder.GetFileExtensions(0)
                                End If
                            Else
                                Directory.CreateDirectory(Path.GetDirectoryName(SyncFilePath))
                                Dim Result As ReturnObject = SafeCopy(SourceFileStream, SyncFilePath)
                                If Not Result.Success Then Throw New Exception(Result.ErrorMessage)
                            End If

                            MyLog.Write(ProcessID, "...successfully added file to sync folder.", Information)
                        End If
                    End If
                Else
                    Throw New Exception("File was being watched but could not determine its codec.")
                End If
            Next

            MyReturnObject = New ReturnObject(True, "", 0)
            MyLog.Write(ProcessID, "Sync file renamed: """ & FilePath.Substring(MyGlobalSyncSettings.SourceDirectory.Length) & """", Information)
        Catch ex As Exception
            MyLog.Write(ProcessID, "Sync file rename failed: """ & FilePath.Substring(MyGlobalSyncSettings.SourceDirectory.Length) & """. Exception: " & ex.Message, Warning)
            MyReturnObject = New ReturnObject(False, ex.Message, 0)
        End Try

        Return MyReturnObject

    End Function

    Private Function TranscodeFile(FileTo As String, SyncSetting As SyncSettings) As ReturnObject

        Dim FileFrom As String = FilePath
        Dim OutputFilePath As String = ""

        Try
            Dim OutputDirectory As String = Path.GetDirectoryName(FileTo)
            OutputFilePath = Path.Combine(OutputDirectory, Path.GetFileNameWithoutExtension(FileTo)) & SyncSetting.Encoder.GetFileExtensions(0)
            Directory.CreateDirectory(OutputDirectory)
        Catch ex As Exception
            Dim MyError As String = ex.Message
            If ex.InnerException IsNot Nothing Then
                MyError &= NewLine & NewLine & ex.InnerException.ToString
            End If
            MyLog.Write(ProcessID, "...transcode failed [1]. Exception: " & MyError, Warning)
            Return New ReturnObject(False, "Transcode failed [1]. Exception: " & MyError)
        End Try

        Try
            Dim ffmpeg As New ProcessStartInfo(MyGlobalSyncSettings.ffmpegPath)
            ffmpeg.CreateNoWindow = True
            ffmpeg.UseShellExecute = False

            Dim FiltersString As String = ""
            If SyncSetting.ReplayGainSetting <> ReplayGainMode.None Then
                FiltersString = " -af volume=replaygain=" & SyncSetting.GetReplayGainSetting().ToLower
            End If

            ffmpeg.Arguments = "-i """ & FileFrom & """ -vn -c:a " & SyncSetting.Encoder.GetProfiles(0).Argument & FiltersString & " -hide_banner """ & OutputFilePath & """"
            'EXAMPLE: libvorbis -aq: 4 = 128 kbps, 5 = 160 kbps, 6 = 192 kbps, 7 = 224 kbps, 8 = 256 kbps

            MyLog.Write(ProcessID, "...ffmpeg arguments: """ & ffmpeg.Arguments & """...", Debug)

            Dim ffmpegProcess As Process = Process.Start(ffmpeg)
            ffmpegProcess.WaitForExit()

            If ffmpegProcess.ExitCode <> 0 Then
                Throw New Exception("ffmpeg exited with an error! (Code: " & ffmpegProcess.ExitCode & ")")
            End If

            MyLog.Write(ProcessID, "...transcode complete...", Debug)
        Catch ex As Exception
            Dim MyError As String = ex.Message
            If ex.InnerException IsNot Nothing Then
                MyError &= NewLine & NewLine & ex.InnerException.ToString
            End If
            MyLog.Write(ProcessID, "...transcode failed [2]. Exception: " & MyError, Warning)
            Return New ReturnObject(False, "Transcode failed [2]. Exception: " & MyError)
        End Try

        Return New ReturnObject(True, "")

    End Function
#End Region

#Region " Safe File Operations "
    Private Shared Function SafeCopy(SourceFilePath As String, SyncFilePath As String) As ReturnObject
        Dim Result As ReturnObject = Nothing

        If File.Exists(SourceFilePath) Then
            Using SourceFileStream As FileStream = WaitForFile(SourceFilePath, FileMode.Open, FileAccess.Read, FileShare.Read, FileTimeout)
                Result = SafeCopy(SourceFileStream, SyncFilePath)
            End Using
        Else
            Result = New ReturnObject(False, "File does not exist: " & SourceFilePath)
        End If

        Return Result
    End Function

    Private Shared Function SafeCopy(SourceFileStream As FileStream, SyncFilePath As String) As ReturnObject

        Dim MyReturnObject As ReturnObject

        Try
            If SourceFileStream Is Nothing Then Throw New Exception("Could not get file system lock on source file.")
            Dim OutputDirectory As String = Path.GetDirectoryName(SyncFilePath)
            Directory.CreateDirectory(OutputDirectory)
            Using NewFile As FileStream = WaitForFile(SyncFilePath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, FileTimeout)
                If Not NewFile Is Nothing Then
                    SourceFileStream.CopyTo(NewFile)
                    MyReturnObject = New ReturnObject(True, "")
                Else
                    MyReturnObject = New ReturnObject(False, "Could not get file system lock on destination file.")
                End If
            End Using
        Catch ex As Exception
            MyReturnObject = New ReturnObject(False, ex.Message)
        End Try

        Return MyReturnObject

    End Function

    Private Shared Function WaitForFile(fullPath As String, mode As FileMode, access As FileAccess, share As FileShare, timeoutSeconds As Int32) As FileStream
        Dim msBetweenTries As Int32 = 500
        Dim numTries As Int32 = CInt(Math.Ceiling(timeoutSeconds / (msBetweenTries / 1000)))

        For count As Integer = 0 To numTries
            Dim fs As FileStream = Nothing

            Try
                fs = New FileStream(fullPath, mode, access, share)

                fs.ReadByte()
                fs.Seek(0, SeekOrigin.Begin)

                Return fs
            Catch generatedExceptionName As IOException
                If fs IsNot Nothing Then
                    fs.Dispose()
                End If
                Thread.Sleep(msBetweenTries)
            End Try
        Next

        Return Nothing
    End Function
#End Region

#Region " File Checks "
    Private Function CheckFileForSync(ByVal FileCodec As Codec, SyncSetting As SyncSettings) As Boolean

        Try
            If CheckFileTags(FileCodec, SyncSetting) Then
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

    Private Function CheckFileCodec(CodecsToCheck As Codec()) As Codec

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

    Private Function CheckFileTags(MyCodec As Codec, SyncSetting As SyncSettings) As Boolean

        Dim TagsObject As ReturnObject = MyCodec.MatchTag(FilePath, SyncSetting.GetWatcherTags)

        If TagsObject.Success Then
            Return CType(TagsObject.MyObject, Boolean)
        Else
            MyLog.Write(ProcessID, "...could not obtain file tags. File: """ & FilePath & """, Codec: " & MyCodec.Name & ", Exception: " & TagsObject.ErrorMessage, Warning)
            Return False
        End If

    End Function
#End Region

End Class