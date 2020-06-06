#Region " Namespaces "
Imports MusicDirectorySyncer.Toolkit
Imports MusicDirectorySyncer.Logger.LogLevel
Imports MusicDirectorySyncer.FileProcessingQueue.FileProcessingInfo.ActionType
Imports System.Threading
#End Region

Class FileProcessingQueue
    Implements IDisposable

    Private TaskQueueMutex As Object
    Private TasksRunningMutex As Object
    Private ReadOnly FileTaskList As New List(Of FileProcessingInfo)
    Private ReadOnly WaitBeforeSlowProcessing_ms As Int32 = 5000     'Repeated file-changed events within this time period cause file processing to restart
    Private Const WaitBeforeFastProcessing_ms As Int32 = 50       'Non-zero to give time for cancellation events to filter through
    Private TasksRunning As Int64 = 0
    Private ReadOnly MaxThreads As Int64 = 2
    Private IsDisposed As Boolean = False

    Public Sub New(WaitBeforeProcessingFiles_ms As Int32, NewMaxThreads As Int64)

        TaskQueueMutex = New Object
        TasksRunningMutex = New Object
        WaitBeforeSlowProcessing_ms = WaitBeforeProcessingFiles_ms
        MaxThreads = NewMaxThreads

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub ThrowIfCancellationRequested(CancelToken As CancellationToken)
        CancelToken.ThrowIfCancellationRequested()
    End Sub

    Public Function AddTask(MyFileProcessingInfo As FileProcessingInfo) As ReturnObject

        Dim TimeSlept_ms As Int32 = 0
        Dim SleepTime_ms As Int32 = 50
        Dim WaitTime_ms = 0

        'Add the task to the queue
        SyncLock TaskQueueMutex
            MyLog.Write(MyFileProcessingInfo.ProcessID, "Adding to FileTaskList: " & MyFileProcessingInfo.FilePath, Debug)
            FileTaskList.Add(MyFileProcessingInfo)
        End SyncLock

        'Cancel all other file tasks if they are accessing the same file as this task
        CancelFileTaskIfAlreadyRunning(MyFileProcessingInfo)

        'Only wait when a file was changed/created/renamed. A deleted file/directory shouldn't require any transcoding so no point waiting
        Select Case MyFileProcessingInfo.ActionToTake
            Case Is = Changed, Created, Renamed
                WaitTime_ms = WaitBeforeSlowProcessing_ms
            Case Is = Deleted, DirectoryDeleted
                WaitTime_ms = WaitBeforeFastProcessing_ms
        End Select

        'Wait a pre-set amount of time until there's a spare thread to use
        MyLog.Write(MyFileProcessingInfo.ProcessID, "Waiting " & WaitTime_ms & " ms before attempting to process file: " & MyFileProcessingInfo.FilePath, Debug)
        Do
            'Do nothing if we are disposing
            If IsDisposed Then
                MyLog.Write(MyFileProcessingInfo.ProcessID, "IsDisposed", Debug)
                Return New ReturnObject(False, "File system watcher has been stopped, aborting pending file processing.")
            End If

            'Cancel this task if we've been told to
            ThrowIfCancellationRequested(MyFileProcessingInfo.CancelState.Token)

            'If we've waited long enough, check if we can start
            If TimeSlept_ms >= WaitTime_ms Then
                'Check number of tasks running to see if we can continue
                SyncLock TasksRunningMutex
                    If TasksRunning < MaxThreads Then
                        TasksRunning += 1
                        MyLog.Write(MyFileProcessingInfo.ProcessID, "Tasks running after increment: " & TasksRunning, Debug)
                        Exit Do
                    End If
                End SyncLock
            End If

            'Wait
            Thread.Sleep(SleepTime_ms)
            TimeSlept_ms += SleepTime_ms
        Loop

        'Start processing the file based on the file event that triggered this task
        Dim Result As ReturnObject = ProcessFile(MyFileProcessingInfo)

        SyncLock TasksRunningMutex
            TasksRunning -= 1
            MyLog.Write(MyFileProcessingInfo.ProcessID, "Tasks running after decrement: " & TasksRunning, Debug)
        End SyncLock

        Return Result

    End Function

    Public Function RemoveTaskFromQueue(FileID As Int32) As Boolean
        Dim TaskRemoved As Boolean = False

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

    Public Function CancelFileTaskIfAlreadyRunning(MyFileProcessingInfo As FileProcessingInfo) As Boolean
        Dim TaskCancelled As Boolean = False

        SyncLock TaskQueueMutex
            MyLog.Write(MyFileProcessingInfo.ProcessID, "Checking if file is already being processed: " & MyFileProcessingInfo.FilePath, Debug)
            For Each OtherFileProcessingInfo As FileProcessingInfo In FileTaskList
                If OtherFileProcessingInfo.ProcessID <> MyFileProcessingInfo.ProcessID AndAlso OtherFileProcessingInfo.FilePath = MyFileProcessingInfo.FilePath Then 'Cancel task
                    MyFileProcessingInfo.CancelState.Cancel()
                    TaskCancelled = True
                    MyLog.Write(MyFileProcessingInfo.ProcessID, "File is already being processed, so cancelling original task", Debug)
                End If
            Next
        End SyncLock

        Return TaskCancelled
    End Function

    Public Sub PrintRunningTasks()
        If MyLog.DebugLevel = Logger.LogLevel.Debug AndAlso CountTasksAlreadyRunning() > 0 Then
            MyLog.Write("[[ REMAINING TASKS: ]]", Debug)
            Dim i As Int32 = 0
            Dim SafeFileTaskList As List(Of FileProcessingInfo)
            SyncLock TaskQueueMutex
                SafeFileTaskList = FileTaskList
            End SyncLock
            For Each MyFileProcessingInfo As FileProcessingInfo In SafeFileTaskList
                i += 1
                MyLog.Write("[[" & i & ": " & MyFileProcessingInfo.FilePath & "]]", Debug)
            Next
        Else
            MyLog.Write("[[ NO REMAINING TASKS ]] ", Debug)
        End If
    End Sub

    Public Function CountTasksAlreadyRunning() As Int32
        Dim Result As Int32 = 0

        SyncLock TaskQueueMutex
            Result = FileTaskList.Count
        End SyncLock

        Return Result
    End Function

    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overloads Sub Dispose(Disposing As Boolean)
        'Ensure this isn't called twice
        If IsDisposed Then Exit Sub

        If (Disposing) Then
            'Wait for current tasks to exit
            Dim SleepTime_ms As Int32 = 50
            Do
                SyncLock TasksRunningMutex
                    If TasksRunning = 0 Then
                        Exit Do
                    End If
                End SyncLock
                Thread.Sleep(SleepTime_ms)
            Loop

            'Clear file task list
            FileTaskList.Clear()

            'Dispose of task queue mutex
            If TaskQueueMutex IsNot Nothing Then
                TaskQueueMutex = Nothing
            End If

            'Dispose of task running mutex
            If TasksRunningMutex IsNot Nothing Then
                TasksRunningMutex = Nothing
            End If
        End If

        IsDisposed = True

    End Sub

    Public Class FileProcessingInfo

        Implements IDisposable

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

        Public Overridable Sub Dispose() Implements IDisposable.Dispose
            If Not FileParser Is Nothing Then
                FileParser.Dispose()
            End If
        End Sub

    End Class

    Private Function ProcessFile(MyFileProcessingInfo As FileProcessingInfo) As ReturnObject

        Dim Result As ReturnObject = Nothing
        Dim SuccessMessage = ""
        Dim FailureMessage = ""

        Select Case MyFileProcessingInfo.ActionToTake
            Case Is = FileProcessingInfo.ActionType.DirectoryDeleted
                MyLog.Write(MyFileProcessingInfo.ProcessID, "Processing directory: " & MyFileProcessingInfo.FilePath, Information)
                Result = MyFileProcessingInfo.FileParser.DeleteDirectoryInSyncDirectory()
                SuccessMessage = "Directory deleted:"
                FailureMessage = "Directory deletion failed:"
            Case Is = FileProcessingInfo.ActionType.Deleted
                MyLog.Write(MyFileProcessingInfo.ProcessID, "Processing file: " & MyFileProcessingInfo.FilePath, Information)
                Result = MyFileProcessingInfo.FileParser.DeleteInSyncDirectory()
                SuccessMessage = "File deleted:"
                FailureMessage = "File deletion failed:"
            Case Is = FileProcessingInfo.ActionType.Renamed
                MyLog.Write(MyFileProcessingInfo.ProcessID, "Processing file: " & MyFileProcessingInfo.FilePath, Information)
                Result = MyFileProcessingInfo.FileParser.RenameInSyncDirectory(MyFileProcessingInfo.OldFilePath)
                SuccessMessage = "File processed:"
                FailureMessage = "File processing failed:"
            Case Is = FileProcessingInfo.ActionType.Created
                MyLog.Write(MyFileProcessingInfo.ProcessID, "Processing file: " & MyFileProcessingInfo.FilePath, Information)
                Result = MyFileProcessingInfo.FileParser.TransferToSyncDirectory()
                SuccessMessage = "File processed:"
                FailureMessage = "File processing failed:"
            Case Is = FileProcessingInfo.ActionType.Changed
                MyLog.Write(MyFileProcessingInfo.ProcessID, "Processing file: " & MyFileProcessingInfo.FilePath, Information)
                Result = MyFileProcessingInfo.FileParser.DeleteInSyncDirectory(True)
                If Result.Success Then Result = MyFileProcessingInfo.FileParser.TransferToSyncDirectory()
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
