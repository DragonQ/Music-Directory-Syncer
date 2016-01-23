Imports Music_Folder_Syncer.Logger.DebugLogLevel
Imports System.Environment


Public Class Logger

    Public Enum DebugLogLevel
        Debug
        Information
        Warning
        Always
    End Enum

    Dim FilePath As String
    Private ReadOnly LoggerSync As New Object


    Public Sub New(ByVal LogFilePath As String)
        FilePath = LogFilePath
    End Sub

    Public Overloads Sub Write(ThreadID As UInt32, Text As String, Optional LogLevel As DebugLogLevel = Always)

        WriteCommon(Text, ThreadID, LogLevel)

    End Sub

    Public Overloads Sub Write(Text As String, Optional LogLevel As DebugLogLevel = Always)

        'If no ThreadID is specified, this is the main thread (0)
        WriteCommon(Text, 0, LogLevel)

    End Sub

    Private Sub WriteCommon(Text As String, ThreadID As UInt32, LogLevel As DebugLogLevel)

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
                LogMarker &= String.Format("[{0:00000}] ", ThreadID)
            End If

            SyncLock LoggerSync
                Dim MyNow As Date = DateTime.Now
                Dim LogText As String = MyNow.ToString & LogMarker & Text & NewLine

                IO.File.AppendAllText(FilePath, LogText)
                Console.WriteLine(LogText)
            End SyncLock
        End If

    End Sub

End Class
