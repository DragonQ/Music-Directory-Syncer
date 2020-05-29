#Region " Namespaces "
Imports MusicDirectorySyncer.Logger.LogLevel
Imports System.Environment
#End Region

Public Class Logger

    Public Enum LogLevel
        Debug
        Information
        Warning
        Fatal
        Always
    End Enum

    Private ReadOnly FilePath As String
    Private ReadOnly LoggerSync As New Object
    Property DebugLevel As LogLevel

    Public Sub New(ByVal LogFilePath As String, ByVal LogDebugLevel As LogLevel)

        FilePath = LogFilePath
        DebugLevel = LogDebugLevel

    End Sub

    Public Shared Function ConvertStringToLogLevel(LogLevelString As String) As LogLevel

        Select Case LogLevelString
            Case Is = "Debug"
                Return Debug
            Case Is = "Information"
                Return Information
            Case Is = "Warning"
                Return Warning
            Case Is = "Fatal"
                Return Fatal
            Case Else
                Return Always
        End Select

    End Function

    Public Shared Function ConvertLogLevelToString(LogLevel As LogLevel) As String

        Select Case LogLevel
            Case Is = Debug
                Return "Debug"
            Case Is = Information
                Return "Information"
            Case Is = Warning
                Return "Warning"
            Case Is = Fatal
                Return "Fatal"
            Case Else
                Return "Always"
        End Select

    End Function

    Public Overloads Sub Write(ThreadID As Int32, Text As String, LogLevel As LogLevel)

        WriteCommon(Text, ThreadID, LogLevel)

    End Sub

    Public Overloads Sub Write(Text As String, LogLevel As LogLevel)

        'If no ThreadID is specified, this is the main thread (0)
        WriteCommon(Text, 0, LogLevel)

    End Sub

    Public Overloads Sub Write(Text As String)

        'If no ThreadID is specified, this is the main thread (0)
        WriteCommon(Text, 0, Always)

    End Sub

    Private Sub WriteCommon(Text As String, ThreadID As Int32, LogLevel As LogLevel)

        If LogLevel >= DebugLevel Then
            Dim LogMarker As String

            Select Case LogLevel
                Case Is = Debug
                    LogMarker = " [Debug]         :   "
                Case Is = Information
                    LogMarker = " [Information]   :   "
                Case Is = Warning
                    LogMarker = " [Warning]       :   "
                Case Is = Fatal
                    LogMarker = " [FATAL]         :   "
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
                Dim LogText As String = MyNow.ToString(EnglishGB) & LogMarker & Text

                IO.File.AppendAllText(FilePath, LogText & NewLine)
#If DEBUG Then
                Console.WriteLine(LogText)
#End If
            End SyncLock
        End If

    End Sub

End Class
