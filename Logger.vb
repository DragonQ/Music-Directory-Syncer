Imports MusicFolderSyncer.Logger.DebugLogLevel
Imports System.Environment


Public Class Logger

    Public Enum DebugLogLevel
        Debug
        Information
        Warning
        Always
    End Enum

    Private ReadOnly FilePath As String
    Private ReadOnly LoggerSync As New Object
    Property DebugLevel As DebugLogLevel

    Public Sub New(ByVal LogFilePath As String, ByVal LogDebugLevel As DebugLogLevel)
        FilePath = LogFilePath
        DebugLevel = LogDebugLevel
    End Sub

    Public Overloads Sub Write(ThreadID As Int32, Text As String, LogLevel As DebugLogLevel)

        WriteCommon(Text, ThreadID, LogLevel)

    End Sub

    Public Overloads Sub Write(Text As String, LogLevel As DebugLogLevel)

        'If no ThreadID is specified, this is the main thread (0)
        WriteCommon(Text, 0, LogLevel)

    End Sub

    Public Overloads Sub Write(Text As String)

        'If no ThreadID is specified, this is the main thread (0)
        WriteCommon(Text, 0, Always)

    End Sub

    Private Sub WriteCommon(Text As String, ThreadID As Int32, LogLevel As DebugLogLevel)

        If LogLevel >= DebugLevel Then
            Dim LogMarker As String

            Select Case LogLevel
                Case Is = Debug
                    LogMarker = " [Debug]         :   "
                Case Is = Information
                    LogMarker = " [Information]   :   "
                Case Is = Warning
                    LogMarker = " [Warning]       :   "
                Case Else
                    LogMarker = "                 :   "
            End Select

            If ThreadID = UInt32.MaxValue Then 'Unknown thread
                LogMarker &= "[?????] "
            Else
                LogMarker &= String.Format(EnglishGB, "[{0:00000}] ", ThreadID)
            End If

            SyncLock LoggerSync
                Dim MyNow As Date = DateTime.Now
                Dim LogText As String = MyNow.ToString(EnglishGB) & LogMarker & Text & NewLine

                IO.File.AppendAllText(FilePath, LogText)
#If DEBUG Then
                Console.WriteLine(LogText)
#End If
            End SyncLock
        End If

    End Sub

End Class
