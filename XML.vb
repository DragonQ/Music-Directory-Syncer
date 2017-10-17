#Region " Namespaces "
Imports MusicFolderSyncer.Toolkit
Imports MusicFolderSyncer.SyncSettings
Imports System.IO
Imports System.Xml
#End Region

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
            Using CodecFileReader As New StringReader(CodecFile)
                Dim CodecsXML As XDocument = XDocument.Load(CodecFileReader)

                'Find all <Codec> elements and assign properties based on known sub-elements
                Dim Codecs = From Codec In CodecsXML.Elements("MusicFolderSyncer").Elements("Codecs").Elements("Codec")
                             Select New With
                    {
                    .Name = Codec.Element("Name").Value,
                    .Type = Codec.Element("Type").Value,
                    .Profiles = Codec.Element("Profiles"),
                    .Extensions = Codec.Element("Extensions")
                    }

                For Each MyCodec In Codecs
                    'Find all <Profile> elements within <Profiles>
                    If MyCodec.Profiles Is Nothing Then
                        Throw New Exception("Codecs file is malformed: missing <Profiles> element.")
                    End If

                    Dim elProfiles = From Profile In MyCodec.Profiles.Elements("Profile")
                                     Select New With
                        {
                        .ProfileName = Profile.Element("Name").Value,
                        .ProfileArgument = Profile.Element("Argument").Value
                        }
                    Dim Profiles As New List(Of Codec.Profile)
                    For Each Profile In elProfiles
                        Profiles.Add(New Codec.Profile(Profile.ProfileName, Profile.ProfileArgument))
                    Next

                    Dim Extensions = From Extension In MyCodec.Extensions.Elements("Extension") Select Extension.Value

                    CodecList.Add(New Codec(MyCodec.Name, MyCodec.Type, Profiles.ToArray, Extensions.ToArray))
                Next
            End Using

            Return CodecList
        Catch ex As Exception
            MyLog.Write("Could not read from Codecs.xml! Error: " & ex.Message, Logger.LogLevel.Warning)
            Return Nothing
        End Try

    End Function

    Public Function ReadSyncSettings(Codecs As List(Of Codec), DefaultSettings As GlobalSyncSettings) As ReturnObject
        Return ReadSettings(Codecs, Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "SyncSettings.xml"), DefaultSettings)
    End Function

    Public Function ReadDefaultSettings(Codecs As List(Of Codec)) As ReturnObject
        Dim DefaultSync As New SyncSettings("", New List(Of Codec), New List(Of Codec.Tag), TranscodeMode.LosslessOnly, Nothing, ReplayGainMode.None)
        Dim DefaultSyncList As New List(Of SyncSettings)
        DefaultSyncList.Add(DefaultSync)
        Return ReadSettings(Codecs, Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "DefaultSettings.xml"), New GlobalSyncSettings(False, "", "", Environment.ProcessorCount, DefaultSyncList, "Information"))
    End Function

    Public Function ReadSettings(Codecs As List(Of Codec), FilePath As String, DefaultSettings As GlobalSyncSettings) As ReturnObject

        Dim NewSyncSettingsList As New List(Of SyncSettings)
        Dim GlobalSyncSettings As New GlobalSyncSettings(DefaultSettings)

        Try
            'If SyncSettings.xml doesn't exist, return nothing
            If Not File.Exists(FilePath) Then Return New ReturnObject(True, "", Nothing)

            'Create StringWriter to store XML text
            Dim SettingsFile As String = File.ReadAllText(FilePath)

            'Read file as XML document
            Using SettingsFileReader As New StringReader(SettingsFile)
                Dim SettingsXML As XDocument = XDocument.Load(SettingsFileReader)

                'Find global settings
                Dim GlobalSettings = From Setting In SettingsXML.Elements("MusicFolderSyncer")
                                     Select New With
                    {
                        .EnableSync = If(Setting.OptionalElement("EnableSync"), Nothing),
                        .SourceDirectory = If(Setting.Element("SourceDirectory"), Nothing),
                        .ffmpegPath = If(Setting.Element("ffmpegPath"), Nothing),
                        .LogLevel = If(Setting.OptionalElement("LogLevel"), Nothing),
                        .MaxThreads = If(Setting.OptionalElement("Threads"), Nothing),
                        .SyncSettings = If(Setting.Elements("SyncSettings"), Nothing)
                    }

                If GlobalSettings.Count > 0 Then
                    If GlobalSettings.Count = 1 Then
                        If Not GlobalSettings(0).EnableSync Is Nothing Then GlobalSyncSettings.SyncIsEnabled = True
                        If Not GlobalSettings(0).SourceDirectory Is Nothing Then GlobalSyncSettings.SourceDirectory = GlobalSettings(0).SourceDirectory.Value
                        If Not GlobalSettings(0).ffmpegPath Is Nothing Then GlobalSyncSettings.ffmpegPath = GlobalSettings(0).ffmpegPath.Value
                        If Not GlobalSettings(0).LogLevel Is Nothing Then GlobalSyncSettings.SetLogLevel(GlobalSettings(0).LogLevel)
                        If Not GlobalSettings(0).MaxThreads Is Nothing Then GlobalSyncSettings.MaxThreads = CInt(GlobalSettings(0).MaxThreads)
                    Else
                        Throw New Exception("Settings file is malformed: too many <MusicFolderSyncer> elements.")
                    End If
                Else
                    Throw New Exception("Settings file is malformed: missing <MusicFolderSyncer> element.")
                End If

                Dim Settings = From Setting In GlobalSettings(0).SyncSettings.Elements("SyncSetting")
                               Select New With
                    {
                        .SyncDirectory = If(Setting.Element("SyncDirectory"), Nothing),
                        .TranscodeSetting = If(Setting.OptionalElement("TranscodeMode"), Nothing),
                        .ReplayGain = If(Setting.OptionalElement("ReplayGainMode"), Nothing),
                        .Encoder = If(Setting.Element("Encoder"), Nothing),
                        .FileTypes = If(Setting.Element("FileTypes"), Nothing),
                        .Tags = If(Setting.Element("Tags"), Nothing)
                    }

                Dim SyncSettingsCount = Settings.Count

                If SyncSettingsCount > 0 Then
                    For Each MySetting In Settings
                        'Apply all default values before searching for this sync's settings
                        Dim NewSyncSettings As New SyncSettings(DefaultSettings.GetSyncSettings(0))

                        If Not MySetting.SyncDirectory Is Nothing Then NewSyncSettings.SyncDirectory = MySetting.SyncDirectory.Value

                        If Not MySetting.ReplayGain Is Nothing Then
                            NewSyncSettings.SetReplayGainSetting(MySetting.ReplayGain)
                        End If

                        If Not MySetting.TranscodeSetting Is Nothing Then
                            NewSyncSettings.SetTranscodeSetting(MySetting.TranscodeSetting)

                            'Find transcoding settings
                            Dim elCodecName As XElement = If(MySetting.Encoder.Element("CodecName"), Nothing)
                            Dim elCodecProfile As XElement = If(MySetting.Encoder.Element("CodecProfile"), Nothing)
                            Dim elExtension As XElement = If(MySetting.Encoder.Element("Extension"), Nothing)

                            If elCodecName Is Nothing Or elCodecProfile Is Nothing Or elExtension Is Nothing Then
                                Throw New Exception("Settings file is malformed: <TranscodeLosslessFiles /> is present but there is no valid <Encoder> elements.")
                            Else
                                Dim CodecName As String = elCodecName.Value
                                Dim CodecProfile As String = elCodecProfile.Value
                                Dim Extension As String = elExtension.Value

                                Dim CodecFound As Boolean = False
                                For Each MyCodec As Codec In Codecs
                                    If CodecName = MyCodec.Name Then
                                        Dim ProfileFound As Boolean = False

                                        For Each MyProfile As Codec.Profile In MyCodec.GetProfiles()
                                            If CodecProfile = MyProfile.Name Then
                                                NewSyncSettings.Encoder = New Codec(MyCodec.Name, MyCodec.TypeString, {MyProfile}, {Extension})
                                                ProfileFound = True
                                                Exit For
                                            End If
                                        Next

                                        If Not ProfileFound Then Throw New Exception("Codec profile not recognised: " & CodecProfile)

                                        CodecFound = True
                                        Exit For
                                    End If
                                Next

                                If Not CodecFound Then
                                    Throw New Exception("Codec not recognised: " & CodecName)
                                End If
                            End If
                        End If

                        'Find all <CodecName> elements within <FileTypes>
                        If MySetting.FileTypes Is Nothing Then
                            Throw New Exception("Settings file is malformed: missing <FileTypes> element.")
                        End If

                        Dim ImportedCodecs = From ImportedCodec In MySetting.FileTypes.Elements("CodecName") Select ImportedCodec.Value
                        If ImportedCodecs Is Nothing Then
                            Throw New Exception("Settings file is malformed: <FileTypes> is present but there is no valid <CodecName> elements.")
                        Else
                            'Add List of codec names to sync to the global filter list
                            Dim WatcherFilterList As New List(Of Codec)
                            If ImportedCodecs.Count > 0 Then
                                For Each ImportedCodec In ImportedCodecs
                                    Dim CodecFound As Boolean = False

                                    For Each MyCodec As Codec In Codecs
                                        If ImportedCodec = MyCodec.Name Then
                                            WatcherFilterList.Add(MyCodec)
                                            CodecFound = True
                                            Exit For
                                        End If
                                    Next

                                    If Not CodecFound Then Throw New Exception("Codec not recognised: " & ImportedCodec)
                                Next
                            Else
                                Throw New Exception("Settings file is malformed: missing <FileTypes> element.")
                            End If
                            NewSyncSettings.SetWatcherCodecs(WatcherFilterList)
                        End If

                        'Find all <Tag> elements within <Tags>
                        If MySetting.Tags Is Nothing Then
                            Throw New Exception("Settings file is malformed: missing <Tags> element.")
                        End If

                        Dim Tags = From Tag In MySetting.Tags.Elements("Tag")
                                   Select New With
                            {
                                .Name = Tag.Element("Name").Value,
                                .Value = If(Tag.OptionalElement("Value"), Nothing)
                            }

                        If Tags Is Nothing Then
                            Throw New Exception("Settings file is malformed: <Tags> is present but there are no valid <Tag> elements.")
                        Else
                            Dim TagList As New List(Of Codec.Tag)

                            For Each Tag In Tags
                                TagList.Add(New Codec.Tag(Tag.Name, Tag.Value))
                            Next

                            NewSyncSettings.SetWatcherTags(TagList)
                        End If

                        NewSyncSettingsList.Add(NewSyncSettings)
                    Next
                Else
                    Throw New Exception("Settings file is malformed: no <SyncSettings> elements.")
                End If
            End Using

            If NewSyncSettingsList.Count > 0 Then
                GlobalSyncSettings.SetSyncSettings(NewSyncSettingsList)
                Return New ReturnObject(True, "", GlobalSyncSettings)
            Else
                Throw New Exception("No sync settings found!")
            End If
        Catch ex As Exception
            Return New ReturnObject(False, ex.Message, Nothing)
        End Try

    End Function

    Public Function SaveSyncSettings(MyGlobalSyncSettings As GlobalSyncSettings) As ReturnObject

        Try
            Dim XMLSettings As XmlWriterSettings = New XmlWriterSettings()
            XMLSettings.Indent = True
            XMLSettings.Encoding = New System.Text.UTF8Encoding
            Dim StringToWrite As String = Nothing

            'Create StringWriter to store XML text
            Using FileData As New StringWriter(EnglishGB)

                'Create XML writer to write all data in XML format
                Dim MyWriter As XmlWriter = XmlWriter.Create(FileData, XMLSettings)
                'Begin document
                MyWriter.WriteStartDocument()
                MyWriter.WriteStartElement("MusicFolderSyncer")

                If MyGlobalSyncSettings.SyncIsEnabled Then
                    MyWriter.WriteStartElement("EnableSync")
                    MyWriter.WriteEndElement()
                End If
                MyWriter.WriteElementString("SourceDirectory", MyGlobalSyncSettings.SourceDirectory)
                MyWriter.WriteElementString("ffmpegPath", MyGlobalSyncSettings.ffmpegPath)
                MyWriter.WriteElementString("LogLevel", MyGlobalSyncSettings.GetLogLevel())
                MyWriter.WriteElementString("Threads", MyGlobalSyncSettings.MaxThreads.ToString(EnglishGB))

                'Write data for each defined sync
                Dim SyncSettings As SyncSettings() = MyGlobalSyncSettings.GetSyncSettings()

                MyWriter.WriteStartElement("SyncSettings")
                If SyncSettings.Count > 0 Then
                    For Each SyncSetting In SyncSettings
                        MyWriter.WriteStartElement("SyncSetting")
                        MyWriter.WriteElementString("SyncDirectory", SyncSetting.SyncDirectory)
                        MyWriter.WriteElementString("ReplayGainMode", SyncSetting.GetReplayGainSetting())

                        'Write data for each codec/file type to be synced
                        MyWriter.WriteStartElement("FileTypes")
                        For Each MyCodec As Codec In SyncSetting.GetWatcherCodecs
                            MyWriter.WriteElementString("CodecName", MyCodec.Name)
                        Next
                        MyWriter.WriteEndElement()

                        'Write data for each file tag to be synced
                        MyWriter.WriteStartElement("Tags")
                        For Each MyTag As Codec.Tag In SyncSetting.GetWatcherTags
                            MyWriter.WriteStartElement("Tag")
                            MyWriter.WriteElementString("Name", MyTag.Name)
                            If Not MyTag.Value Is Nothing Then MyWriter.WriteElementString("Value", MyTag.Value)
                            MyWriter.WriteEndElement()
                        Next
                        MyWriter.WriteEndElement()

                        MyWriter.WriteElementString("TranscodeMode", SyncSetting.GetTranscodeSetting())
                        If SyncSetting.TranscodeSetting <> TranscodeMode.None Then
                            MyWriter.WriteStartElement("Encoder")
                            MyWriter.WriteElementString("CodecName", SyncSetting.Encoder.Name)
                            MyWriter.WriteElementString("CodecProfile", SyncSetting.Encoder.GetProfiles(0).Name)
                            MyWriter.WriteElementString("Extension", SyncSetting.Encoder.GetFileExtensions(0))
                            MyWriter.WriteEndElement()
                        End If

                        MyWriter.WriteEndElement()
                    Next
                End If
                MyWriter.WriteEndElement()

                'End document
                MyWriter.WriteEndElement()
                MyWriter.WriteEndDocument()
                MyWriter.Flush()

                'Write XML text to string
                StringToWrite = FileData.ToString.Replace("<?xml version=""1.0"" encoding=""utf-16""?>",
                                                                    "<?xml version=""1.0"" encoding=""utf-8""?>")
            End Using

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
