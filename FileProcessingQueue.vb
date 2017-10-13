#Region " Namespaces "
Imports MusicFolderSyncer.Toolkit
Imports MusicFolderSyncer.Logger.LogLevel
Imports System.Threading
#End Region

Class FileProcessingQueue
    Implements IDisposable

    Private TaskQueueMutex As New Mutex()
    Private FileTaskList As New List(Of TaskDescriptor)
    Private WaitBeforeProcessingFiles_ms As Int32 = 5000    'Repeated events within this time period cause file processing to restart
    Private TasksRunning As Int64 = 0
    Private MaxThreads As Int64 = 2
    Private IsDisposing As Boolean = False

    Public Sub New(NewWaitBeforeProcessingFiles_ms As Int32, NewMaxThreads As Int64)

        WaitBeforeProcessingFiles_ms = NewWaitBeforeProcessingFiles_ms
        MaxThreads = NewMaxThreads

    End Sub

    Public Function AddTask(MyFileProcessingInfo As FileProcessingInfo) As ReturnObject

        Dim TimeSlept_ms As Int32 = 0
        Dim SleepTime_ms As Int32 = 50

        Dim NewTaskDescriptor As New FileProcessingQueue.TaskDescriptor(MyFileProcessingInfo.ProcessID, MyFileProcessingInfo.FilePath, MyFileProcessingInfo.CancelState)

        TaskQueueMutex.WaitOne()
        FileTaskList.Add(NewTaskDescriptor)
        TaskQueueMutex.ReleaseMutex()

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
        If TimeSlept_ms > WaitBeforeProcessingFiles_ms Then
            MyLog.Write(MyFileProcessingInfo.ProcessID, "Already waited " & TimeSlept_ms & " ms before attempting to process file: " & MyFileProcessingInfo.FilePath, Debug)
        Else
            MyLog.Write(MyFileProcessingInfo.ProcessID, "Waiting " & WaitBeforeProcessingFiles_ms & " ms before attempting to process file: " & MyFileProcessingInfo.FilePath, Debug)
            Thread.Sleep(WaitBeforeProcessingFiles_ms - TimeSlept_ms)
        End If

        Dim Result As ReturnObject = FileChangedAction(MyFileProcessingInfo, TimeSlept_ms)
        Interlocked.Decrement(TasksRunning)
        Return Result

    End Function

    Public Function RemoveTaskFromQueue(FileID As Int32) As Boolean
        Dim TaskRemoved As Boolean = False
        Dim TaskIndexes As New List(Of Int32)

        TaskQueueMutex.WaitOne()

        For i As Int32 = 0 To FileTaskList.Count - 1
            If FileTaskList(i).FileID = FileID Then 'Cancel task
                FileTaskList.RemoveAt(i)
                TaskRemoved = True
                Exit For
            End If
        Next

        MyLog.Write(FileID, "Tasks still running: " & CountTasksAlreadyRunning(), Information)

        TaskQueueMutex.ReleaseMutex()

        Return TaskRemoved
    End Function

    Public Function CancelFileTaskIfAlreadyRunning(FilePath As String) As Boolean
        Dim TaskCancelled As Boolean = False

        TaskQueueMutex.WaitOne()

        For Each MyTaskDescriptor As TaskDescriptor In FileTaskList
            If MyTaskDescriptor.FilePath = FilePath Then 'Cancel task
                MyTaskDescriptor.CancelToken.Cancel()
                TaskCancelled = True
            End If
        Next

        TaskQueueMutex.ReleaseMutex()

        Return TaskCancelled
    End Function

    Public Sub PrintRunningTasks()
        If MyLog.DebugLevel = Logger.LogLevel.Debug AndAlso CountTasksAlreadyRunning() > 0 Then
            MyLog.Write(0, "[[[ REMAINING TASKS: ]]]", Debug)
            Dim i As Int32 = 0
            Dim FileTaskList As List(Of TaskDescriptor) = GetFileTaskList()
            For Each MyTaskDescriptor As TaskDescriptor In FileTaskList
                i += 1
                MyLog.Write(0, i & ": " & MyTaskDescriptor.FilePath, Debug)
            Next
        Else
            MyLog.Write(0, "[[[ NO REMAINING TASKS ]]] ", Debug)
        End If
    End Sub

    Public Function CountTasksAlreadyRunning() As Int32
        Dim Result As Int32 = 0

        TaskQueueMutex.WaitOne()
        Result = FileTaskList.Count
        TaskQueueMutex.ReleaseMutex()

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
            TaskQueueMutex.Close()
        End If

        GC.SuppressFinalize(Me)

    End Sub

    Public Class FileProcessingInfo
        ReadOnly Property ProcessID As Int32
        ReadOnly Property FilePath As String
        ReadOnly Property FileParser As FileParser
        ReadOnly Property IsNewFile As Boolean
        Property CancelState As CancellationTokenSource
        Public Property ResultMessages As String()

        Public Sub New(NewProcessID As Int32, NewFilePath As String, NewFileParser As FileParser, NewFile As Boolean, ByRef NewCancelState As CancellationTokenSource)
            ProcessID = NewProcessID
            FilePath = NewFilePath
            FileParser = NewFileParser
            IsNewFile = NewFile
            CancelState = NewCancelState
        End Sub

    End Class

    Private Class TaskDescriptor
        Public ReadOnly Property FileID As Int32
        Public ReadOnly Property FilePath As String
        Public Property CancelToken As CancellationTokenSource

        Public Sub New(NewFileID As Int32, NewFilePath As String, NewCancellationTokenSource As CancellationTokenSource)
            FileID = NewFileID
            FilePath = NewFilePath
            CancelToken = NewCancellationTokenSource
        End Sub
    End Class

    Private Function GetFileTaskList() As List(Of TaskDescriptor)
        TaskQueueMutex.WaitOne()
        Return FileTaskList
        TaskQueueMutex.ReleaseMutex()
    End Function

    Private Function FileChangedAction(MyFileProcessingInfo As FileProcessingInfo, TimeAlreadySlept_ms As Int32) As ReturnObject

        Dim Result As ReturnObject = Nothing

        '=================================== DEBUG CODE BELOW:
        'Dim threadid As Int32 = Thread.CurrentThread.ManagedThreadId
        'MyLog.Write(MyFileProcessingInfo.ProcessID, threadid & " ------ Source file changed: " & MyFileProcessingInfo.FilePath, Information)
        '===================================

        'If MyFileProcessingInfo.CancelState.Token.IsCancellationRequested() Then
        '    MyLog.Write(MyFileProcessingInfo.ProcessID, "File processing cancelled before attempting to process file: " & MyFileProcessingInfo.FilePath, Warning)
        '    Return New ReturnObject(False, "Task cancelled", MyFileProcessingInfo)
        'End If

        MyFileProcessingInfo.CancelState.Token.ThrowIfCancellationRequested()

        'No other tasks are running using this file, so we are free to continue
        MyLog.Write(MyFileProcessingInfo.ProcessID, "Processing file: " & MyFileProcessingInfo.FilePath, Information)
        If Not MyFileProcessingInfo.IsNewFile Then
            Result = MyFileProcessingInfo.FileParser.DeleteInSyncFolder()
        End If

        If MyFileProcessingInfo.IsNewFile OrElse Result.Success Then
            Result = MyFileProcessingInfo.FileParser.TransferToSyncFolder()
            If Result.Success Then
                MyFileProcessingInfo.ResultMessages = {"File processed:", MyFileProcessingInfo.FilePath.Substring(UserGlobalSyncSettings.SourceDirectory.Length)}
                Return New ReturnObject(True, "", MyFileProcessingInfo)
            End If
        End If

        'Failure case
        Return Result

    End Function
End Class
