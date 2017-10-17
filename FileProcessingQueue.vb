#Region " Namespaces "
Imports MusicFolderSyncer.Toolkit
Imports MusicFolderSyncer.Logger.LogLevel
Imports MusicFolderSyncer.FileProcessingQueue.FileProcessingInfo.ActionType
Imports System.Threading
#End Region

Class FileProcessingQueue
    Implements IDisposable

    Private TaskQueueMutex As New Object
    Private FileTaskList As New List(Of FileProcessingInfo)
    Private WaitBeforeSlowProcessing_ms As Int32 = 5000     'Repeated file-changed events within this time period cause file processing to restart
    Private WaitBeforeFastProcessing_ms As Int32 = 50       'Non-zero to give time for cancellation events to filter through
    Private TasksRunning As Int64 = 0
    Private MaxThreads As Int64 = 2
    Private IsDisposing As Boolean = False

    Public Sub New(WaitBeforeProcessingFiles_ms As Int32, NewMaxThreads As Int64)

        WaitBeforeSlowProcessing_ms = WaitBeforeProcessingFiles_ms
        MaxThreads = NewMaxThreads

    End Sub

    Public Function AddTask(MyFileProcessingInfo As FileProcessingInfo) As ReturnObject

        Dim TimeSlept_ms As Int32 = 0
        Dim SleepTime_ms As Int32 = 50
        Dim WaitTime_ms = 0

        SyncLock TaskQueueMutex
            FileTaskList.Add(MyFileProcessingInfo)
        End SyncLock

        'Only wait when a file was changed/created. A deleted or renamed file shouldn't require any transcoding so no point waiting
        Select Case MyFileProcessingInfo.ActionToTake
            Case Is = Changed, Created
                WaitTime_ms = WaitBeforeSlowProcessing_ms
            Case Is = Renamed, Deleted, DirectoryDeleted
                WaitTime_ms = WaitBeforeFastProcessing_ms
        End Select

        'Wait until there's a spare thread to use
        Do Until Interlocked.Read(TasksRunning) < MaxThreads
            If IsDisposing Then
                Return New ReturnObject(False, "File system watcher has been stopped, aborting pending file processing.", Nothing)
            End If
            TimeSlept_ms += SleepTime_ms
            Thread.Sleep(SleepTime_ms)
        Loop
        Interlocked.Increment(TasksRunning)

        'Wait a pre-set amount of time and then cancel this task if we've been told to
        If WaitTime_ms > 0 Then
            If TimeSlept_ms >= WaitTime_ms Then
                MyLog.Write(MyFileProcessingInfo.ProcessID, "Already waited " & TimeSlept_ms & " ms before attempting to process file: " & MyFileProcessingInfo.FilePath, Debug)
            Else
                MyLog.Write(MyFileProcessingInfo.ProcessID, "Waiting " & WaitTime_ms & " ms before attempting to process file: " & MyFileProcessingInfo.FilePath, Debug)
                Thread.Sleep(WaitTime_ms - TimeSlept_ms)
            End If
        End If

        'Start processing the file based on the file event that triggered this task
        Dim Result As ReturnObject = ProcessFile(MyFileProcessingInfo)
        Interlocked.Decrement(TasksRunning)
        Return Result

    End Function

    Public Function RemoveTaskFromQueue(FileID As Int32) As Boolean
        Dim TaskRemoved As Boolean = False
        Dim TaskIndexes As New List(Of Int32)

        SyncLock TaskQueueMutex

            For i As Int32 = 0 To FileTaskList.Count - 1
                If FileTaskList(i).ProcessID = FileID Then 'Cancel task
                    FileTaskList.RemoveAt(i)
                    TaskRemoved = True
                    Exit For
                End If
            Next

        End SyncLock

        Return TaskRemoved
    End Function

    Public Function CancelFileTaskIfAlreadyRunning(FilePath As String) As Boolean
        Dim TaskCancelled As Boolean = False

        SyncLock TaskQueueMutex

            For Each MyFileProcessingInfo As FileProcessingInfo In FileTaskList
                If MyFileProcessingInfo.FilePath = FilePath Then 'Cancel task
                    MyFileProcessingInfo.CancelState.Cancel()
                    TaskCancelled = True
                End If
            Next

        End SyncLock

        Return TaskCancelled
    End Function

    Public Sub PrintRunningTasks()
        If MyLog.DebugLevel = Logger.LogLevel.Debug AndAlso CountTasksAlreadyRunning() > 0 Then
            MyLog.Write(0, "[[ REMAINING TASKS: ]]", Debug)
            Dim i As Int32 = 0
            Dim FileTaskList As List(Of FileProcessingInfo) = GetFileTaskList()
            For Each MyFileProcessingInfo As FileProcessingInfo In FileTaskList
                i += 1
                MyLog.Write(0, "[[" & i & ": " & MyFileProcessingInfo.FilePath & "]]", Debug)
            Next
        Else
            MyLog.Write(0, "[[ NO REMAINING TASKS ]] ", Debug)
        End If
    End Sub

    Public Function CountTasksAlreadyRunning() As Int32
        Dim Result As Int32 = 0

        SyncLock TaskQueueMutex
            Result = FileTaskList.Count
        End SyncLock

        Return Result
    End Function

    Public Sub Dispose() Implements IDisposable.Dispose

        'Ensure this isn't called twice
        If IsDisposing Then Exit Sub
        IsDisposing = True

        'Wait for current tasks to exit
        Do Until Interlocked.Read(TasksRunning) = 0
            Thread.Sleep(50)
        Loop

        'Dispose of task queue mutex
        If TaskQueueMutex IsNot Nothing Then
            TaskQueueMutex = Nothing
        End If

        GC.SuppressFinalize(Me)

    End Sub

    Public Class FileProcessingInfo
        ReadOnly Property ProcessID As Int32
        ReadOnly Property FilePath As String
        ReadOnly Property OldFilePath As String
        ReadOnly Property FileParser As FileParser
        ReadOnly Property ActionToTake As ActionType
        Property CancelState As CancellationTokenSource
        Public Property ResultMessages As String()

        Enum ActionType
            Changed
            Created
            Renamed
            Deleted
            DirectoryDeleted
        End Enum

        Public Sub New(NewProcessID As Int32, NewFilePath As String, NewFileParser As FileParser, ChangeType As ActionType, ByRef NewCancelState As CancellationTokenSource, Optional OriginalFilePath As String = "")
            ProcessID = NewProcessID
            FilePath = NewFilePath
            FileParser = NewFileParser
            CancelState = NewCancelState
            OldFilePath = OriginalFilePath
            ActionToTake = ChangeType
        End Sub

    End Class

    Private Function GetFileTaskList() As List(Of FileProcessingInfo)
        Dim Result As List(Of FileProcessingInfo) = Nothing

        SyncLock TaskQueueMutex
            Result = FileTaskList
        End SyncLock

        Return Result
    End Function

    Private Function ProcessFile(MyFileProcessingInfo As FileProcessingInfo) As ReturnObject

        Dim Result As ReturnObject = Nothing
        Dim SuccessMessage = ""
        Dim FailureMessage = ""

        MyFileProcessingInfo.CancelState.Token.ThrowIfCancellationRequested()

        'No other tasks are running using this file, so we are free to continue
        Select Case MyFileProcessingInfo.ActionToTake
            Case Is = FileProcessingInfo.ActionType.DirectoryDeleted
                MyLog.Write(MyFileProcessingInfo.ProcessID, "Processing directory: " & MyFileProcessingInfo.FilePath, Information)
                Result = MyFileProcessingInfo.FileParser.DeleteDirectoryInSyncFolder()
                SuccessMessage = "Directory deleted:"
                FailureMessage = "Directory deletion failed:"
            Case Is = FileProcessingInfo.ActionType.Deleted
                MyLog.Write(MyFileProcessingInfo.ProcessID, "Processing file: " & MyFileProcessingInfo.FilePath, Information)
                Result = MyFileProcessingInfo.FileParser.DeleteInSyncFolder()
                SuccessMessage = "File deleted:"
                FailureMessage = "File deletion failed:"
            Case Is = FileProcessingInfo.ActionType.Renamed
                MyLog.Write(MyFileProcessingInfo.ProcessID, "Processing file: " & MyFileProcessingInfo.FilePath, Information)
                Result = MyFileProcessingInfo.FileParser.RenameInSyncFolder(MyFileProcessingInfo.OldFilePath)
                SuccessMessage = "File processed:"
                FailureMessage = "File processing failed:"
            Case Is = FileProcessingInfo.ActionType.Created
                MyLog.Write(MyFileProcessingInfo.ProcessID, "Processing file: " & MyFileProcessingInfo.FilePath, Information)
                Result = MyFileProcessingInfo.FileParser.TransferToSyncFolder()
                SuccessMessage = "File processed:"
                FailureMessage = "File processing failed:"
            Case Is = FileProcessingInfo.ActionType.Changed
                MyLog.Write(MyFileProcessingInfo.ProcessID, "Processing file: " & MyFileProcessingInfo.FilePath, Information)
                Result = MyFileProcessingInfo.FileParser.DeleteInSyncFolder()
                If Result.Success Then Result = MyFileProcessingInfo.FileParser.TransferToSyncFolder()
                SuccessMessage = "File processed:"
                FailureMessage = "File processing failed:"
        End Select

        If Result.Success Then
            MyFileProcessingInfo.ResultMessages = {SuccessMessage, MyFileProcessingInfo.FilePath.Substring(UserGlobalSyncSettings.SourceDirectory.Length)}
            Return New ReturnObject(True, "", MyFileProcessingInfo)
        Else
            MyFileProcessingInfo.ResultMessages = {FailureMessage, MyFileProcessingInfo.FilePath.Substring(UserGlobalSyncSettings.SourceDirectory.Length)}
            Return New ReturnObject(False, Result.ErrorMessage, MyFileProcessingInfo)
        End If

    End Function
End Class
