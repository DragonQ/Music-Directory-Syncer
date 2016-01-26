Imports System.IO
Imports System.Xml

Module XML

    ''' <summary>
    ''' This function makes it easy to search for optional XML elements within an
    ''' XDocument object. If <paramref name="ElementName"/> exists within
    ''' <paramref name="ActionElement"/> , its value is returned. Otherwise, nothing
    ''' is returned.
    ''' </summary>
    ''' <param name="ActionElement">XElement in which to search.</param>
    ''' <param name="ElementName">XElement to search for.</param>
    <System.Runtime.CompilerServices.Extension> _
    Public Function OptionalElement(ActionElement As XElement, ElementName As String) As String
        Dim Element = ActionElement.Element(ElementName)
        Return If((Element IsNot Nothing), Element.Value, Nothing)
    End Function

    Public Function ReadCodecs() As List(Of Codec)

        Dim CodecList As New List(Of Codec)

        Try
            'Create StringWriter to store XML text
            Dim CodecFile As String = File.ReadAllText(Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "Codecs.xml"))

            'Read file as XML document
            Dim CodecsXML As XDocument = XDocument.Load(New StringReader(CodecFile))

            'Find all <Codec> elements and assign properties based on known sub-elements
            Dim Codecs = From Codec In CodecsXML.Elements("MusicFolderSyncer").Elements("Codecs").Elements("Codec")
                         Select New With
                {
                .Name = Codec.Element("Name").Value,
                .Type = Codec.Element("Type").Value,
                .Extensions = From Extension In Codec.Descendants("Extension")
                }

            'Find all <Codec> elements and assign properties based on known sub-elements
            Dim Codecs2 = From Codec In CodecsXML.Elements("MusicFolderSyncer").Elements("Codecs").Elements("Codec")
                          Select New With
                {
                .Name = Codec.Element("Name").Value,
                .Profile = From Profile In Codec.Descendants("Profile")
                           Select New With
                    {
                    .ProfileName = Profile.Element("Name").Value,
                    .ProfileArgument = Profile.Element("Argument").Value
                    }
                }

            For Each MyCodec In Codecs
                Dim Extensions As New List(Of String)
                Dim Profiles As New List(Of Codec.Profile)

                For Each MyExtension In MyCodec.Extensions
                    Extensions.Add(MyExtension.Value)
                Next
                For Each MyCodec2 In Codecs2
                    If MyCodec2.Name = MyCodec.Name Then
                        For Each MyProfile In MyCodec2.Profile
                            Profiles.Add(New Codec.Profile(MyProfile.ProfileName, MyProfile.ProfileArgument))
                        Next
                        Exit For
                    End If
                Next
                CodecList.Add(New Codec(MyCodec.Name, MyCodec.Type, Profiles.ToArray, Extensions.ToArray))
            Next

            Return CodecList
        Catch ex As Exception
            MyLog.Write("Could not read from Codecs.xml! Error: " & ex.Message, Logger.DebugLogLevel.Warning)
            Return Nothing
        End Try

    End Function

    Public Function ReadSyncSettings(Codecs As List(Of Codec), DefaultSettings As SyncSettings) As ReturnObject
        Return ReadSettings(Codecs, Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "SyncSettings.xml"), DefaultSettings)
    End Function

    Public Function ReadDefaultSettings(Codecs As List(Of Codec)) As ReturnObject
        Return ReadSettings(Codecs, Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "DefaultSettings.xml"))
    End Function

    Public Function ReadSettings(Codecs As List(Of Codec), FilePath As String, Optional DefaultSettings As SyncSettings = Nothing) As ReturnObject

        Dim WatcherFilterList As New List(Of Codec)
        Dim TagList As New List(Of Codec.Tag)
        Dim NewSyncSettings As SyncSettings

        Try
            'If SyncSettings.xml doesn't exist, return nothing
            If Not File.Exists(FilePath) Then Return New ReturnObject(True, "", Nothing)

            'Apply all default values before searching for this sync's settings
            If Not DefaultSettings Is Nothing Then
                NewSyncSettings = New SyncSettings(DefaultSettings)
            Else
                NewSyncSettings = New SyncSettings(False, "", "", New List(Of Codec), New List(Of Codec.Tag), False, Nothing, System.Environment.ProcessorCount, "")
            End If

            'Create StringWriter to store XML text
            Dim SettingsFile As String = File.ReadAllText(FilePath)

            'Read file as XML document
            Dim SettingsXML As XDocument = XDocument.Load(New StringReader(SettingsFile))

            'Find all settings
            Dim Settings = From Setting In SettingsXML.Elements("MusicFolderSyncer")
                           Select New With
                {
                    .SourceDirectory = If(Setting.Element("SourceDirectory"), Nothing),
                    .SyncDirectory = If(Setting.Element("SyncDirectory"), Nothing),
                    .MaxThreads = If(Setting.OptionalElement("Threads"), Nothing),
                    .EnableSync = If(Setting.OptionalElement("EnableSync"), Nothing),
                    .TranscodeLosslessFiles = If(Setting.OptionalElement("TranscodeLosslessFiles"), Nothing),
                    .ffmpegPath = If(Setting.Element("ffmpegPath"), Nothing)
                }

            If Settings.Count > 0 Then
                If Settings.Count = 1 Then
                    If Not Settings(0).EnableSync Is Nothing Then NewSyncSettings.SyncIsEnabled = True
                    If Not Settings(0).SourceDirectory Is Nothing Then NewSyncSettings.SourceDirectory = Settings(0).SourceDirectory.Value
                    If Not Settings(0).SyncDirectory Is Nothing Then NewSyncSettings.SyncDirectory = Settings(0).SyncDirectory.Value
                    If Not Settings(0).MaxThreads Is Nothing Then NewSyncSettings.MaxThreads = CInt(Settings(0).MaxThreads)
                    If Not Settings(0).ffmpegPath Is Nothing Then NewSyncSettings.ffmpegPath = Settings(0).ffmpegPath.Value

                    If Not Settings(0).TranscodeLosslessFiles Is Nothing Then
                        NewSyncSettings.TranscodeLosslessFiles = True

                        'Find transcoding settings
                        Dim TranscodeSettings = From Setting In SettingsXML.Elements("MusicFolderSyncer").Elements("Encoder")
                                                Select New With
                            {
                                .CodecName = Setting.Element("CodecName"),
                                .CodecProfile = Setting.Element("CodecProfile"),
                                .Extension = Setting.Element("Extension")
                            }

                        If TranscodeSettings.Count > 0 Then
                            If TranscodeSettings.Count = 1 Then
                                Dim CodecFound As Boolean = False

                                For Each MyCodec As Codec In Codecs
                                    If TranscodeSettings(0).CodecName.Value = MyCodec.Name Then

                                        Dim ProfileFound As Boolean = False

                                        For Each MyProfile As Codec.Profile In MyCodec.Profiles
                                            If TranscodeSettings(0).CodecProfile.Value = MyProfile.Name Then
                                                NewSyncSettings.Encoder = New Codec(MyCodec.Name, MyCodec.GetTypeString, {MyProfile},
                                                                           {TranscodeSettings(0).Extension.Value})
                                                ProfileFound = True
                                                Exit For
                                            End If
                                        Next

                                        If Not ProfileFound Then
                                            If Not ProfileFound Then Throw New Exception("Codec profile not recognised: " & TranscodeSettings(0).CodecProfile.Value)
                                        End If

                                        CodecFound = True
                                        Exit For
                                    End If
                                Next

                                If CodecFound Then

                                Else
                                    Throw New Exception("Codec not recognised: " & TranscodeSettings(0).CodecName.Value)
                                End If
                            Else
                                Throw New Exception("Settings file is malformed: too many <Encoder> elements.")
                            End If
                        Else
                            Throw New Exception("Settings file is malformed: <TranscodeLosslessFiles /> is present but there is no <Encoder> element.")
                        End If

                    End If

                Else
                    Throw New Exception("Settings file is malformed: too many <MusicFolderSyncer> elements.")
                End If
            Else
                Throw New Exception("Settings file is malformed: missing <MusicFolderSyncer> element.")
            End If


            'Find all <CodecName> elements within <FileTypes>
            Dim ImportedCodecs = From ImportedCodec In SettingsXML.Elements("MusicFolderSyncer").Elements("FileTypes").Elements("CodecName")
                                 Select New With
                {
                .Name = ImportedCodec.Value
                }

            'Add list of codec names to sync to the global filter list
            If ImportedCodecs.Count > 0 Then
                For Each ImportedCodec In ImportedCodecs
                    Dim CodecFound As Boolean = False

                    For Each MyCodec As Codec In Codecs
                        If ImportedCodec.Name = MyCodec.Name Then
                            WatcherFilterList.Add(MyCodec)
                            CodecFound = True
                            Exit For
                        End If
                    Next

                    If Not CodecFound Then Throw New Exception("Codec not recognised: " & ImportedCodec.Name)
                Next
            Else
                Throw New Exception("Settings file is malformed: missing <FileTypes> element.")
            End If

            NewSyncSettings.SetWatcherCodecs(WatcherFilterList)

            Dim Tags = From Tag In SettingsXML.Elements("MusicFolderSyncer").Elements("Tags").Elements("Tag")
                       Select New With
                {
                    .Name = Tag.Element("Name").Value,
                    .Value = If(Tag.OptionalElement("Value"), Nothing)
                }

            For Each Tag In Tags
                TagList.Add(New Codec.Tag(Tag.Name, Tag.Value))
            Next

            NewSyncSettings.SetWatcherTags(TagList)

            Return New ReturnObject(True, "", NewSyncSettings)
        Catch ex As Exception
            Return New ReturnObject(False, ex.Message, Nothing)
        End Try

    End Function

    Public Function SaveSyncSettings() As ReturnObject

        Try
            Dim XMLSettings As XmlWriterSettings = New XmlWriterSettings()
            XMLSettings.Indent = True
            XMLSettings.Encoding = New System.Text.UTF8Encoding

            'Create StringWriter to store XML text
            Dim FileData As New StringWriter

            'Create XML writer to write all data in XML format
            Using MyWriter As XmlWriter = XmlWriter.Create(FileData, XMLSettings)
                'Begin document
                MyWriter.WriteStartDocument()
                MyWriter.WriteStartElement("MusicFolderSyncer")

                If MySyncSettings.SyncIsEnabled Then
                    MyWriter.WriteStartElement("EnableSync")
                    MyWriter.WriteEndElement()
                End If
                MyWriter.WriteElementString("Threads", MySyncSettings.MaxThreads.ToString)
                MyWriter.WriteElementString("SourceDirectory", MySyncSettings.SourceDirectory)
                MyWriter.WriteElementString("SyncDirectory", MySyncSettings.SyncDirectory)
                MyWriter.WriteElementString("ffmpegPath", MySyncSettings.ffmpegPath)

                'Write data for each codec/file type to be synced
                MyWriter.WriteStartElement("FileTypes")
                For Each MyCodec As Codec In MySyncSettings.GetWatcherCodecs
                    MyWriter.WriteElementString("CodecName", MyCodec.Name)
                Next
                MyWriter.WriteEndElement()

                'Write data for each file tag to be synced
                MyWriter.WriteStartElement("Tags")
                For Each MyTag As Codec.Tag In MySyncSettings.GetWatcherTags
                    MyWriter.WriteStartElement("Tag")
                    MyWriter.WriteElementString("Name", MyTag.Name)
                    If Not MyTag.Value Is Nothing Then MyWriter.WriteElementString("Value", MyTag.Value)
                    MyWriter.WriteEndElement()
                Next
                MyWriter.WriteEndElement()

                If MySyncSettings.TranscodeLosslessFiles Then
                    MyWriter.WriteStartElement("TranscodeLosslessFiles")
                    MyWriter.WriteEndElement()
                    MyWriter.WriteStartElement("Encoder")
                    MyWriter.WriteElementString("CodecName", MySyncSettings.Encoder.Name)
                    MyWriter.WriteElementString("CodecProfile", MySyncSettings.Encoder.Profiles(0).Name)
                    MyWriter.WriteElementString("Extension", MySyncSettings.Encoder.FileExtensions(0))
                    MyWriter.WriteEndElement()
                End If

                'End document
                MyWriter.WriteEndElement()
                MyWriter.WriteEndDocument()
                MyWriter.Flush()
            End Using

            'Write XML text to string
            Dim StringToWrite As String = FileData.ToString.Replace("<?xml version=""1.0"" encoding=""utf-16""?>",
                                                                    "<?xml version=""1.0"" encoding=""utf-8""?>")

            If StringToWrite Is Nothing OrElse StringToWrite = "" Then 'Something's gone wrong
                Throw New Exception("Failed to create data file.")
            Else
                Dim FilePath As String = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "SyncSettings.xml")
                File.WriteAllText(FilePath, StringToWrite)
            End If

            Return New ReturnObject(True, "", Nothing)
        Catch ex As Exception
            Return New ReturnObject(False, ex.Message, Nothing)
        End Try

    End Function
End Module
