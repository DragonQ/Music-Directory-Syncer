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
            End Using

            Return CodecList
        Catch ex As Exception
            MyLog.Write("Could not read from Codecs.xml! Error: " & ex.Message, Logger.DebugLogLevel.Warning)
            Return Nothing
        End Try

    End Function

    Public Function ReadSyncSettings(Codecs As List(Of Codec), DefaultSettings As GlobalSyncSettings) As ReturnObject
        Return ReadSettings(Codecs, Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "SyncSettings.xml"), DefaultSettings)
    End Function

    Public Function ReadDefaultSettings(Codecs As List(Of Codec)) As ReturnObject
        Dim DefaultSync As New SyncSettings("", New List(Of Codec), New List(Of Codec.Tag), False, Nothing, Environment.ProcessorCount, ReplayGainMode.None)
        Dim DefaultSyncList As New List(Of SyncSettings)
        DefaultSyncList.Add(DefaultSync)
        Return ReadSettings(Codecs, Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "DefaultSettings.xml"), New GlobalSyncSettings(False, "", "", DefaultSyncList))
    End Function

    Public Function ReadSettings(Codecs As List(Of Codec), FilePath As String, DefaultSettings As GlobalSyncSettings) As ReturnObject

        Dim WatcherFilterList As New List(Of Codec)
        Dim TagList As New List(Of Codec.Tag)
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
                        .ffmpegPath = If(Setting.Element("ffmpegPath"), Nothing)
                    }

                If GlobalSettings.Count > 0 Then
                    If GlobalSettings.Count = 1 Then
                        If Not GlobalSettings(0).EnableSync Is Nothing Then GlobalSyncSettings.SyncIsEnabled = True
                        If Not GlobalSettings(0).SourceDirectory Is Nothing Then GlobalSyncSettings.SourceDirectory = GlobalSettings(0).SourceDirectory.Value
                        If Not GlobalSettings(0).ffmpegPath Is Nothing Then GlobalSyncSettings.ffmpegPath = GlobalSettings(0).ffmpegPath.Value
                    Else
                        Throw New Exception("Settings file is malformed: too many <MusicFolderSyncer> elements.")
                    End If
                Else
                    Throw New Exception("Settings file is malformed: missing <MusicFolderSyncer> element.")
                End If

                'Find all sync settings
                Dim Settings = From Setting In SettingsXML.Elements("MusicFolderSyncer").Elements("SyncSettings").Elements("SyncSetting")
                               Select New With
                    {
                        .SourceDirectory = If(Setting.Element("SourceDirectory"), Nothing),
                        .SyncDirectory = If(Setting.Element("SyncDirectory"), Nothing),
                        .MaxThreads = If(Setting.OptionalElement("Threads"), Nothing),
                        .TranscodeLosslessFiles = If(Setting.OptionalElement("TranscodeLosslessFiles"), Nothing),
                        .ReplayGain = If(Setting.OptionalElement("ReplayGainMode"), Nothing)
                    }

                If Settings.Count > 0 Then
                    For Each MySetting In Settings
                        'Apply all default values before searching for this sync's settings
                        Dim NewSyncSettings As New SyncSettings(DefaultSettings.GetSyncSettings(0))

                        If Not MySetting.SyncDirectory Is Nothing Then NewSyncSettings.SyncDirectory = MySetting.SyncDirectory.Value
                        If Not MySetting.MaxThreads Is Nothing Then NewSyncSettings.MaxThreads = CInt(MySetting.MaxThreads)

                        If Not MySetting.ReplayGain Is Nothing Then
                            Select Case MySetting.ReplayGain
                                Case Is = "Album"
                                    NewSyncSettings.ReplayGain = ReplayGainMode.Album
                                Case Is = "Track"
                                    NewSyncSettings.ReplayGain = ReplayGainMode.Track
                                Case Else
                                    NewSyncSettings.ReplayGain = ReplayGainMode.None
                            End Select
                        End If

                        If Not MySetting.TranscodeLosslessFiles Is Nothing Then
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

                                            For Each MyProfile As Codec.Profile In MyCodec.GetProfiles()
                                                If TranscodeSettings(0).CodecProfile.Value = MyProfile.Name Then
                                                    NewSyncSettings.Encoder = New Codec(MyCodec.Name, MyCodec.TypeString, {MyProfile},
                                                                               {TranscodeSettings(0).Extension.Value})
                                                    ProfileFound = True
                                                    Exit For
                                                End If
                                            Next

                                            If Not ProfileFound Then Throw New Exception("Codec profile not recognised: " & TranscodeSettings(0).CodecProfile.Value)

                                            CodecFound = True
                                            Exit For
                                        End If
                                    Next

                                    If Not CodecFound Then
                                        Throw New Exception("Codec not recognised: " & TranscodeSettings(0).CodecName.Value)
                                    End If
                                Else
                                    Throw New Exception("Settings file is malformed: too many <Encoder> elements.")
                                End If
                            Else
                                Throw New Exception("Settings file is malformed: <TranscodeLosslessFiles /> is present but there is no <Encoder> element.")
                            End If

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

                'Write data for each defined sync
                Dim SyncSettings As SyncSettings() = MyGlobalSyncSettings.GetSyncSettings()

                MyWriter.WriteStartElement("SyncSettings")
                If SyncSettings.Count > 0 Then
                    For Each SyncSetting In SyncSettings
                        MyWriter.WriteStartElement("SyncSetting")
                        MyWriter.WriteElementString("Threads", SyncSetting.MaxThreads.ToString(EnglishGB))
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

                        If SyncSetting.TranscodeLosslessFiles Then
                            MyWriter.WriteStartElement("TranscodeLosslessFiles")
                            MyWriter.WriteEndElement()
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
