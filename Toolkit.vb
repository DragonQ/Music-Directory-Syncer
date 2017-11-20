Imports System.IO
Imports Microsoft.WindowsAPICodePack.Dialogs

'Useful generic tools
Namespace Toolkit

#Region " ComboBox/TickBox Item "
    'Custom content item for combo boxes or tickbox lists
    Public Class Item
        Property Name As String
        Property Value As Object

        Public Sub New(MyName As String, MyValue As Object)
            If Not MyName Is Nothing Then Name = MyName
            If Not MyValue Is Nothing Then Value = MyValue
        End Sub

        Public Sub New(MyItem As Item) ' Cloning
            If Not MyItem Is Nothing Then
                Name = MyItem.Name
                Value = MyItem.Value
            End If
        End Sub

        Public Overrides Function ToString() As String
            'Generates the text shown in the control
            Return Name
        End Function
    End Class
#End Region

#Region " Generic Return Object "
    Public Class ReturnObject
        Property Success As Boolean
        Property ErrorMessage As String
        Property MyObject As Object

        Public Sub New(MySuccess As Boolean, MyErrorMessage As String, ReturnObject As Object)
            Success = MySuccess
            ErrorMessage = MyErrorMessage
            MyObject = ReturnObject
        End Sub

        Public Sub New(MySuccess As Boolean, MyErrorMessage As String)
            Success = MySuccess
            ErrorMessage = MyErrorMessage
            MyObject = Nothing
        End Sub
    End Class
#End Region

#Region " File/Directory Browsers "
    Public Class Browsers
        Public Shared Function CreateDirectoryBrowser(StartingDirectory As String, Title As String) As ReturnObject

            Using SelectDirectoryDialog = New CommonOpenFileDialog()
                With SelectDirectoryDialog
                    .Title = Title
                    .IsFolderPicker = True
                    .AddToMostRecentlyUsedList = False
                    .AllowNonFileSystemItems = True
                    .DefaultDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    If Directory.Exists(StartingDirectory) Then
                        .InitialDirectory = StartingDirectory
                    Else
                        .InitialDirectory = .DefaultDirectory
                    End If
                    .EnsureFileExists = False
                    .EnsurePathExists = True
                    .EnsureReadOnly = False
                    .EnsureValidNames = True
                    .Multiselect = False
                    .ShowPlacesList = True
                End With

                If SelectDirectoryDialog.ShowDialog() = CommonFileDialogResult.Ok Then
                    Return New ReturnObject(True, "", SelectDirectoryDialog.FileName)
                Else
                    Return New ReturnObject(False, "")
                End If
            End Using

        End Function

        Public Shared Function CreateFileBrowser(DefaultPath As String, Title As String, Filters As CommonFileDialogFilter()) As ReturnObject

            Using SelectFileDialog = New CommonOpenFileDialog()
                With SelectFileDialog
                    .Title = Title
                    .IsFolderPicker = False
                    .AddToMostRecentlyUsedList = False
                    .AllowNonFileSystemItems = True

                    .DefaultDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
                    If File.Exists(DefaultPath) Then
                        .InitialDirectory = Path.GetDirectoryName(DefaultPath)
                    Else
                        .InitialDirectory = .DefaultDirectory
                    End If

                    .EnsureFileExists = True
                    .EnsurePathExists = True
                    .EnsureReadOnly = False
                    .EnsureValidNames = True
                    .Multiselect = False
                    .ShowPlacesList = True
                    For Each Filter As CommonFileDialogFilter In Filters
                        .Filters.Add(Filter)
                    Next
                End With

                If SelectFileDialog.ShowDialog() = CommonFileDialogResult.Ok Then
                    Return New ReturnObject(True, "", SelectFileDialog.FileName)
                Else
                    Return New ReturnObject(False, "")
                End If
            End Using

        End Function

        Public Shared Function CreateFileBrowser_ffmpeg(DefaultPath As String) As ReturnObject
            Return CreateFileBrowser(DefaultPath, "Select ffmpeg.exe Path", {New CommonFileDialogFilter("ffmpeg Executable", "*.exe")})
        End Function
    End Class
#End Region

End Namespace