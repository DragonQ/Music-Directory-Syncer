#Region " Namespaces "
Imports System.IO
#End Region

Public Class MyFileSystemWatcher
    Inherits FileSystemWatcher
    Implements IDisposable

#Region " Private Members "
    ' This Dictionary keeps the track of when an event occured last for a particular file
    Private _lastFileEvent As Dictionary(Of String, DateTime)
    ' Interval in Millisecond
    Private _interval As Integer
    'Timespan created when interval is set
    Private _recentTimeSpan As TimeSpan
#End Region

#Region " Constructors "
    ' Constructors delegate to the base class constructors and call private InitializeMember method
    Public Sub New()
        MyBase.New()
        InitializeMembers()
    End Sub

    Public Sub New(Path As String)
        MyBase.New(Path)
        InitializeMembers()
    End Sub

    Public Sub New(Path As String, Filter As String)
        MyBase.New(Path, Filter)
        InitializeMembers()
    End Sub
#End Region

#Region " Events "
    ' These events hide the events from the base class. 
    ' We want to raise these events appropriately and we do not want the 
    ' users of this class subscribing to these events of the base class accidentally
    Public Shadows Event Changed As FileSystemEventHandler
    Public Shadows Event Created As FileSystemEventHandler
    Public Shadows Event Deleted As FileSystemEventHandler
    Public Shadows Event Renamed As RenamedEventHandler
#End Region

#Region " Protected Methods "
    ' Protected Methods to raise the Events for this class
    Protected Overridable Shadows Sub OnChanged(e As FileSystemEventArgs)
        'MyLog.Write(1000, "@@@@@@@@@@ CHANGED RAISED", Logger.LogLevel.Warning)
        RaiseEvent Changed(Me, e)
    End Sub

    Protected Overridable Shadows Sub OnCreated(e As FileSystemEventArgs)
        'MyLog.Write(1000, "@@@@@@@@@@ CREATED RAISED", Logger.LogLevel.Warning)
        RaiseEvent Created(Me, e)
    End Sub

    Protected Overridable Shadows Sub OnDeleted(e As FileSystemEventArgs)
        RaiseEvent Deleted(Me, e)
    End Sub

    Protected Overridable Shadows Sub OnRenamed(e As RenamedEventArgs)
        RaiseEvent Renamed(Me, e)
    End Sub
#End Region

#Region " Private Methods "

    ''' <summary>
    ''' This Method Initializes the private members.
    ''' Interval is set to its default value of 100 millisecond
    ''' FilterRecentEvents is set to true, _lastFileEvent dictionary is initialized
    ''' We subscribe to the base class events.
    ''' </summary>
    Private Sub InitializeMembers()
        Interval = 100
        FilterRecentEvents = True
        _lastFileEvent = New Dictionary(Of String, DateTime)()

        AddHandler MyBase.Created, New FileSystemEventHandler(AddressOf OnCreated)
        AddHandler MyBase.Changed, New FileSystemEventHandler(AddressOf OnChanged)
        AddHandler MyBase.Deleted, New FileSystemEventHandler(AddressOf OnDeleted)
        AddHandler MyBase.Renamed, New RenamedEventHandler(AddressOf OnRenamed)
    End Sub

    ''' <summary>
    ''' This method searches the dictionary to find out when the last event occured 
    ''' for a particular file. If that event occured within the specified timespan
    ''' it returns true, else false
    ''' </summary>
    ''' <param name="FileName">The filename to be checked</param>
    ''' <returns>True if an event has occured within the specified interval, False otherwise</returns>
    Private Function HasAnotherFileEventOccuredRecently(FileName As String) As Boolean
        Dim retVal As Boolean = False

        ' Check dictionary only if user wants to filter recent events otherwise return Value stays False
        If FilterRecentEvents Then
            If _lastFileEvent.ContainsKey(FileName) Then
                ' If dictionary contains the filename, check how much time has elapsed
                ' since the last event occured. If the timespan is less that the 
                ' specified interval, set return value to true 
                ' and store current datetime in dictionary for this file
                Dim lastEventTime As DateTime = _lastFileEvent(FileName)
                Dim currentTime As DateTime = DateTime.Now
                Dim timeSinceLastEvent As TimeSpan = currentTime - lastEventTime
                'MyLog.Write(1000, ">>>>>> " & FileName & " timeSinceLastEvent: " & timeSinceLastEvent.TotalMilliseconds & " ms", Logger.LogLevel.Warning)
                retVal = timeSinceLastEvent < _recentTimeSpan
                _lastFileEvent(FileName) = currentTime
            Else
                ' If dictionary does not contain the filename, 
                ' no event has occured in past for this file, so set return value to false
                ' and annd filename alongwith current datetime to the dictionary
                _lastFileEvent.Add(FileName, DateTime.Now)
                'MyLog.Write(1000, ">>>>>>" & FileName & " no previous event", Logger.LogLevel.Warning)
                retVal = False
            End If
        End If

        Return retVal
    End Function

#Region "FileSystemWatcher EventHandlers"
    ' Base class Event Handlers. Check if an event has occured recently and call method
    ' to raise appropriate event only if no recent event is detected
    Private Shadows Sub OnChanged(sender As Object, e As FileSystemEventArgs)
        'MyLog.Write(1000, "@@@@@@@@ CHANGED NOTICED", Logger.LogLevel.Warning)
        If Not HasAnotherFileEventOccuredRecently(e.FullPath) Then
            Me.OnChanged(e)
        End If
    End Sub

    Private Shadows Sub OnCreated(sender As Object, e As FileSystemEventArgs)
        'MyLog.Write(1000, "@@@@@@@@ CREATED NOTICED", Logger.LogLevel.Warning)
        If Not HasAnotherFileEventOccuredRecently(e.FullPath) Then
            Me.OnCreated(e)
        End If
    End Sub

    Private Shadows Sub OnDeleted(sender As Object, e As FileSystemEventArgs)
        If Not HasAnotherFileEventOccuredRecently(e.FullPath) Then
            Me.OnDeleted(e)
        End If
    End Sub

    Private Shadows Sub OnRenamed(sender As Object, e As RenamedEventArgs)
        If Not HasAnotherFileEventOccuredRecently(e.OldFullPath) Then
            Me.OnRenamed(e)
        End If
    End Sub
#End Region
#End Region

#Region " Public Properties "

    ''' <summary>
    ''' Interval, in milliseconds, within which events are considered "recent"
    ''' </summary>
    Public Property Interval As Integer
        Get
            Return _interval
        End Get
        Set
            _interval = Value
            ' Set timespan based on the value passed
            _recentTimeSpan = New TimeSpan(0, 0, 0, 0, Value)
        End Set
    End Property

    ''' <summary>
    ''' Allows user to set whether to filter recent events. If this is set a false,
    ''' this class behaves like System.IO.FileSystemWatcher class
    ''' </summary>
    Public Property FilterRecentEvents As Boolean
#End Region

#Region " IDisposable Members "

    Protected Overrides Sub Dispose(disposing As Boolean)
        MyBase.Dispose(disposing)
    End Sub

#End Region

End Class