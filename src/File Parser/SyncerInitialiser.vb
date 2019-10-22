#Region " Namespaces "
Imports MusicFolderSyncer.Toolkit
Imports MusicFolderSyncer.Logger.LogLevel
Imports System.IO
Imports System.Environment
Imports System.Threading
Imports System.Security.AccessControl
Imports System.Security.Principal
Imports Opulos.Core.IO
#End Region

Public Class SyncerInitialiser

    Public WithEvents SyncBackgroundWorker As New BackgroundWorker
    Private ThreadsCompleted As Int64 = 0
    Private SyncFolderSize As Int64 = 0
    Private MyGlobalSyncSettings As GlobalSyncSettings
    ReadOnly MySyncSettings As SyncSettings()
    ReadOnly CallbackUpdateMilliseconds As Int32 = 500  'Update UI twice a second by default

    Public Sub New(NewGlobalSyncSettings As GlobalSyncSettings, NewCallbackUpdateMilliseconds As Int32)
        'This is used where we want to re-apply every single sync in one sweep

        'Copy in user settings
        MyGlobalSyncSettings = NewGlobalSyncSettings
        MySyncSettings = MyGlobalSyncSettings.GetSyncSettings()
        CallbackUpdateMilliseconds = NewCallbackUpdateMilliseconds

    End Sub

    Public Sub New(NewGlobalSyncSettings As GlobalSyncSettings, NewSyncSettings As SyncSettings, NewCallbackUpdateMilliseconds As Int32)
        'This is used when we want to add a new sync

        'Copy in user settings
        MyGlobalSyncSettings = NewGlobalSyncSettings
        MySyncSettings = {NewSyncSettings}
        CallbackUpdateMilliseconds = NewCallbackUpdateMilliseconds

    End Sub

    Public Sub AddProgressCallback(CallbackFunction As ProgressChangedEventHandler)
        AddHandler SyncBackgroundWorker.ProgressChanged, CallbackFunction
    End Sub

    Public Sub AddCompletionCallback(CallbackFunction As RunWorkerCompletedEventHandler)
        AddHandler SyncBackgroundWorker.RunWorkerCompleted, CallbackFunction
    End Sub

    Public Sub InitialiseSync()

        'Set properties And events for the SyncBackgroundWorker
        SyncBackgroundWorker.WorkerReportsProgress = True
        SyncBackgroundWorker.WorkerSupportsCancellation = True

        'Begin sync on background thread
        AddHandler SyncBackgroundWorker.DoWork, AddressOf SyncFolder
        SyncBackgroundWorker.RunWorkerAsync()

    End Sub

    Private Sub SyncFolder(sender As Object, e As DoWorkEventArgs)

        Dim FolderPath As String = MyGlobalSyncSettings.SourceDirectory
        Dim MyFiles As IEnumerable(Of FastFileInfo) = Nothing

        Using CancelTokenSource As New CancellationTokenSource()
            MyLog.Write("Scanning files in source directory.", Information)

            Try
                MyFiles = FastFileInfo.EnumerateFiles(FolderPath, "*.*", SearchOption.AllDirectories)
            Catch ex As Exception
                MyLog.Write("Failed to grab file list from source directory. Exception: " & ex.Message, Warning)
                e.Result = New ReturnObject(False, ex.Message)
                Exit Sub
            End Try

            If MyFiles Is Nothing Then
                MyLog.Write("Failed to grab file list from source directory. Exception: EnumerateFiles returned nothing.", Warning)
                e.Result = New ReturnObject(False, "EnumerateFiles returned nothing.")
                Exit Sub
            End If

            Dim ThreadsStarted As UInt32 = 0
            Dim FileID As Int32 = 0
            Dim One As UInt32 = 1

            'Delete existing sync folder if necessary
            For Each Sync As SyncSettings In MySyncSettings
                Try
                    MyLog.Write("Deleting sync folder: " & Sync.SyncDirectory, Information)
                    Directory.Delete(Sync.SyncDirectory, True)
                Catch ex As DirectoryNotFoundException
                    'Do nothing
                Catch ex As Exception
                    Dim MyError As String = ex.Message
                    If ex.InnerException IsNot Nothing Then
                        MyError &= NewLine & NewLine & ex.InnerException.ToString
                    End If
                    MyLog.Write("Failed to delete existing files in sync folder [1]. Exception: " & MyError, Warning)
                    e.Result = New ReturnObject(False, ex.Message)
                    Exit Sub
                End Try
            Next

            'Create sync folder
            For Each Sync As SyncSettings In MySyncSettings
                Try
                    MyLog.Write("Creating sync folder: " & Sync.SyncDirectory, Information)
                    Directory.CreateDirectory(Sync.SyncDirectory, MyDirectoryPermissions)
                Catch ex As Exception
                    Dim MyError As String = ex.Message
                    If ex.InnerException IsNot Nothing Then
                        MyError &= NewLine & NewLine & ex.InnerException.ToString
                    End If
                    MyLog.Write("Failed to delete existing files in sync folder [2]. Exception: " & MyError, Warning)
                    e.Result = New ReturnObject(False, ex.Message)
                    Exit Sub
                End Try
            Next

            MyLog.Write("Starting file processing threads.", Information)
            SyncFolderSize = 0

            Try
                '==============================================================================================
                '============== My own custom way of handling threads, replaced with ThreadPool ===============
                '==============================================================================================
                'Dim ThreadsToRun As Int32 = MyFiles.Count
                'For Each MyFile As FileData In MyFiles
                '    Do
                '        Dim MyThreadsRunning As Int64 = Interlocked.Read(ThreadsRunning)
                '        SyncBackgroundWorker.ReportProgress(CInt(Interlocked.Read(ThreadsCompleted) / ThreadsToRun * 100), MyThreadsRunning)

                '        If MyThreadsRunning < MySyncSettings.MaxThreads Then
                '            Interlocked.Increment(ThreadsRunning)

                '            Dim NewThread As New Thread(New ParameterizedThreadStart(AddressOf TransferToSyncFolderDelegate))
                '            NewThread.IsBackground = True
                '            Dim InputObjects As Object() = {MyFile.Path, CodecsToCheck}

                '            NewThread.Start(InputObjects)

                '            ThreadsStarted += One
                '            MyLog.Write("[" & NewThread.ManagedThreadId & "] Processing file: """ & MyFile.Path & """...", Debug)
                '            Exit Do
                '        Else
                '            Thread.Sleep(50)
                '        End If
                '    Loop
                'Next

                ''All threads have been started, now wait for them all to finish
                'Dim CheckPointThreadsCompleted As Int64 = Interlocked.Read(ThreadsCompleted)

                'Do
                '    Dim ThreadsCompletedSoFar As Int64 = Interlocked.Read(ThreadsCompleted)

                '    If ThreadsCompletedSoFar >= ThreadsStarted Then
                '        SyncBackgroundWorker.ReportProgress(100, 0)
                '        MyLog.Write("All files processed. " & ThreadsStarted & " files synced.", Information)
                '        Exit Do
                '    Else
                '        If ThreadsCompletedSoFar > CheckPointThreadsCompleted Then
                '           SyncBackgroundWorker.ReportProgress(CInt(ThreadsCompletedSoFar / ThreadsStarted * 100), Interlocked.Read(ThreadsRunning))
                '        End If
                '        Thread.Sleep(50)
                '    End If
                'Loop

                '==============================================================================================
                '==============================================================================================

                ThreadPool.SetMinThreads(MyGlobalSyncSettings.MaxThreads, MyGlobalSyncSettings.MaxThreads)
                ThreadPool.SetMaxThreads(MyGlobalSyncSettings.MaxThreads, MyGlobalSyncSettings.MaxThreads)

                For Each MyFile As FastFileInfo In MyFiles
                    If SyncBackgroundWorker.CancellationPending Then
                        CancelTokenSource.Cancel()
                        e.Cancel = True
                        e.Result = New ReturnObject(False, "Sync was cancelled.")
                        Exit Sub
                    End If

                    If FileID = MaxFileID Then
                        FileID = 1
                    Else
                        FileID += 1
                    End If

                    ThreadsStarted += One
                    Dim InputObjects As Object() = {FileID, MyFile.FullName, MySyncSettings, MyDirectoryPermissions, CancelTokenSource.Token}
                    ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf TransferToSyncFolderDelegate), InputObjects)
                Next

                If ThreadsStarted < 1 Then
                    MyLog.Write("Failed to grab file list from source directory. Exception: EnumerateFiles returned no files.", Warning)
                    Throw New Exception("No files found in source directory.")
                End If

                'All threads have been started, now wait for them all to finish
                Dim CheckPointThreadsCompleted As Int64 = Interlocked.Read(ThreadsCompleted)
                Dim ThreadsRemaining As Int64 = ThreadsStarted - CheckPointThreadsCompleted
                SyncBackgroundWorker.ReportProgress(CInt(CheckPointThreadsCompleted / ThreadsStarted * 100),
                                                                {ThreadsRemaining, CheckPointThreadsCompleted})

                Do
                    If SyncBackgroundWorker.CancellationPending Then
                        CancelTokenSource.Cancel()
                        e.Cancel = True
                        e.Result = New ReturnObject(False, "Sync was cancelled.")
                        Exit Sub
                    End If

                    Thread.Sleep(CallbackUpdateMilliseconds)

                    Dim ThreadsCompletedSoFar As Int64 = Interlocked.Read(ThreadsCompleted)

                    If ThreadsCompletedSoFar > CheckPointThreadsCompleted Then 'At least one thread has finished since the last check
                        If ThreadsCompletedSoFar >= ThreadsStarted Then 'All files have been processed, so end background worker
                            MyLog.Write("All files processed. " & ThreadsStarted & " files synced.", Information)
                            Exit Do
                        Else
                            ThreadsRemaining = ThreadsStarted - ThreadsCompletedSoFar
                            SyncBackgroundWorker.ReportProgress(CInt(ThreadsCompletedSoFar / ThreadsStarted * 100),
                                                                {ThreadsRemaining, ThreadsCompletedSoFar})
                        End If
                    End If
                Loop

                '==============================================================================================

                e.Result = New ReturnObject(True, Nothing, SyncFolderSize)
            Catch ex As Exception
                MyLog.Write("Failed to complete sync. Exception: " & ex.Message, Warning)
                e.Result = New ReturnObject(False, ex.Message)
            End Try
        End Using

    End Sub

    Private Sub TransferToSyncFolderDelegate(ByVal Input As Object)

        Try
            Dim InputObjects As Object() = CType(Input, Object())
            Dim CancelToken As CancellationToken = CType(InputObjects(4), CancellationToken)

            If Not CancelToken.IsCancellationRequested Then
                Dim ProcessID As Int32 = CType(InputObjects(0), Int32)
                Dim FilePath As String = CType(InputObjects(1), String)
                Dim NewSyncSettings As SyncSettings() = CType(InputObjects(2), SyncSettings())
                Dim AccessPermissions As DirectorySecurity = CType(InputObjects(3), DirectorySecurity)
                Dim TransferResult As ReturnObject

                Using MyFileParser As New FileParser(MyGlobalSyncSettings, ProcessID, FilePath, AccessPermissions, NewSyncSettings)
                    TransferResult = MyFileParser.TransferToSyncFolder()
                End Using
                If TransferResult.Success Then
                    Dim NewSize As Int64 = CType(TransferResult.MyObject, Int64)

                    If NewSize > 0 Then
                        Interlocked.Add(SyncFolderSize, NewSize)
                    End If
                Else
                    MyLog.Write(ProcessID, "File could not be processed. Error: " & TransferResult.ErrorMessage, Warning)
                End If
            End If
        Catch ex As Exception
            MyLog.Write(Int32.MaxValue, "File could not be processed. Malformed input to TransferToSyncFolderDelegate.", Warning)
        Finally
            Interlocked.Increment(ThreadsCompleted)
            'Interlocked.Decrement(ThreadsRunning)
        End Try

    End Sub

End Class
