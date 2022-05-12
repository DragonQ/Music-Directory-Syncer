#Region " Namespaces "
Imports MusicDirectorySyncer.Toolkit
Imports TagLib
Imports System.Globalization.CultureInfo
#End Region

Public Class Codec

    'Implements INotifyPropertyChanged

    ReadOnly Property Name As String
    Property CompressionType As CodecType
    Private ReadOnly Profiles As Profile()
    Private ReadOnly FileExtensions As String()
    Property IsEnabled As Boolean = True

    Enum CodecType
        Lossless
        Lossy
    End Enum

#Region " New "
    Public Sub New(MyName As String, MyType As String, MyProfiles As Profile(), Extensions As String())

        Name = MyName
        Profiles = MyProfiles
        FileExtensions = Extensions
        CompressionType = ConvertTypeStringToType(MyType)

    End Sub

    Public Sub New(MyCodec As Codec, MyProfile As Profile)
        If MyCodec IsNot Nothing Then
            Name = MyCodec.Name
            FileExtensions = MyCodec.FileExtensions
            CompressionType = MyCodec.CompressionType
        End If
        If MyProfile IsNot Nothing Then Profiles = {MyProfile}
    End Sub
#End Region

    Public Function GetProfiles() As Profile()
        Return Profiles
    End Function

    Public Sub SetType(NewType As CodecType)
        CompressionType = NewType
    End Sub

    Public Function GetFileExtensions() As String()
        Return FileExtensions
    End Function

    Private Shared Function ConvertTypeStringToType(TypeString As String) As CodecType

        Select Case TypeString
            Case Is = "Lossless"
                Return CodecType.Lossless
            Case Is = "Lossy"
                Return CodecType.Lossy
            Case Else
                Throw New Exception("Codec type """ & TypeString & """ not recognised.")
        End Select

    End Function

    Public ReadOnly Property TypeString() As String
        Get
            Select Case CompressionType
                Case Is = CodecType.Lossless
                    Return "Lossless"
                Case Is = CodecType.Lossy
                    Return "Lossy"
                Case Else
                    Return Nothing
            End Select
        End Get
    End Property

    Public Overridable Function MatchTag(FilePath As String, Tags As Tag()) As ReturnObject

        Select Case Me.Name
            Case Is = "FLAC", "OGG"
                Return OggCodec.MatchTag(FilePath, Tags, Me.Name)
            Case Is = "WMA Lossless", "WMA"
                Return WMACodec.MatchTag(FilePath, Tags)
            Case Is = "MP3"
                Return MP3Codec.MatchTag(FilePath, Tags)
            Case Is = "AAC"
                Return AACCodec.MatchTag(FilePath, Tags)
            Case Is = "WavPack"
                Return WavPackCodec.MatchTag(FilePath, Tags)
            Case Else
                Return New ReturnObject(False, "Codec not recognised: " & Me.Name)
        End Select

    End Function

    Public Class Profile
        ReadOnly Property Name As String
        'Property Type As ProfileType
        ReadOnly Property Argument As String

        'Enum ProfileType
        '    CBR
        '    VBR
        'End Enum

        Public Sub New(MyName As String, MyArgument As String)

            Name = MyName
            'Type = MyType
            Argument = MyArgument

        End Sub
    End Class

    Public Class Tag
        Property Name As String
        Property Value As String

        Public Sub New(MyName As String, MyValue As String)

            Name = MyName
            Value = MyValue

        End Sub

        Public Sub New(MyName As String)

            Name = MyName
            Value = Nothing

        End Sub
    End Class

    Private Class Mpeg4TestFile
        Inherits Mpeg4.File
        Public Sub New(path As String)

            MyBase.New(path)
        End Sub

        Public Shadows ReadOnly Property UdtaBoxes() As List(Of Mpeg4.IsoUserDataBox)
            Get
                Return MyBase.UdtaBoxes
            End Get
        End Property
    End Class

    Private Class OggCodec

        Private Sub New()
            ' Unused...
        End Sub

        Public Shared Function MatchTag(FilePath As String, Tags As Tag(), CodecName As String) As ReturnObject

            If Tags IsNot Nothing Then
                Try
                    If Tags.Length = 0 Then
                        'Tags weren't specified, so always match every file
                        Return New ReturnObject(True, "", True)
                    Else
                        'Find tags within OGG/FLAC file
                        Dim MyFile As TagLib.File
                        If CodecName = "FLAC" Then
                            MyFile = New Flac.File(FilePath)
                        ElseIf CodecName = "OGG" Then
                            MyFile = New Ogg.File(FilePath)
                        Else
                            Throw New Exception("Cannot find tags because codec name """ & CodecName & """ is not recognised.")
                        End If

                        Using MyFile
                            Dim Xiph As Ogg.XiphComment = CType(MyFile.GetTag(TagTypes.Xiph, False), Ogg.XiphComment)

                            If Xiph Is Nothing Then
                                Throw New Exception("OGG/FLAC tags not found.")
                            Else
                                'Search for each requested tag
                                For Each MyTag As Tag In Tags
                                    Dim Results As String() = Xiph.GetField(MyTag.Name)

                                    If Results IsNot Nothing AndAlso Results.Length > 0 Then 'Tag we're looking for is present, so continue
                                        'Value matches or wasn't requested, so return true
                                        If MyTag.Value Is Nothing OrElse MyTag.Value = "" OrElse
                                            MyTag.Value.ToUpper(InvariantCulture) = Results(0).Trim.ToUpper(InvariantCulture) Then
                                            Return New ReturnObject(True, "", True)
                                        End If
                                    End If
                                Next

                                'If none of the tags were found, return false
                                Return New ReturnObject(True, "", False)
                            End If
                        End Using

                        Return New ReturnObject(True, "", False)
                    End If
                Catch ex As Exception
                    Return New ReturnObject(False, ex.Message)
                End Try
            Else
                Return New ReturnObject(False, "Tags is nothing")
            End If

        End Function

    End Class

    Private Class WavPackCodec

        Private Sub New()
            ' Unused...
        End Sub

        Public Shared Function MatchTag(FilePath As String, Tags As Tag()) As ReturnObject

            If Tags IsNot Nothing Then
                Try
                    If Tags.Length = 0 Then
                        'Tags weren't specified, so always match every file
                        Return New ReturnObject(True, "", True)
                    Else
                        'Find tags within WV file
                        Using WavPackFile As TagLib.File = TagLib.WavPack.File.Create(FilePath)
                            Dim MyApe As Ape.Tag = CType(WavPackFile.GetTag(TagTypes.Ape, False), Ape.Tag)

                            If MyApe Is Nothing Then
                                Throw New Exception("WavPack tags not found.")
                            Else
                                'Search for each requested tag
                                For Each MyTag As Tag In Tags
                                    If MyApe.HasItem(MyTag.Name) Then 'Tag we're looking for is present, so continue
                                        'Value matches or wasn't requested, so return true
                                        If MyTag.Value Is Nothing OrElse MyTag.Value = "" OrElse MyTag.Value.ToUpper(InvariantCulture) = MyApe.GetItem(MyTag.Name).ToString.Trim.ToUpper(InvariantCulture) Then
                                            Return New ReturnObject(True, "", True)
                                        End If
                                    End If
                                Next

                                'If none of the tags were found, return false
                                Return New ReturnObject(True, "", False)
                            End If
                        End Using

                        Return New ReturnObject(True, "", False)
                    End If
                Catch ex As Exception
                    Return New ReturnObject(False, ex.Message)
                End Try
            Else
                Return New ReturnObject(False, "Tags is nothing")
            End If

        End Function

    End Class

    Private Class WMACodec

        Private Sub New()
            ' Unused...
        End Sub

        Public Shared Function MatchTag(FilePath As String, Tags As Tag()) As ReturnObject

            If Tags IsNot Nothing Then
                Try
                    If Tags.Length = 0 Then
                        'Tags weren't specified, so always match every file
                        Return New ReturnObject(True, "", True)
                    Else
                        'Find tags within WMA file
                        Using WMAFile As TagLib.File = TagLib.File.Create(FilePath)
                            Dim ASF As Asf.Tag = CType(WMAFile.GetTag(TagTypes.Asf, False), Asf.Tag)

                            If ASF Is Nothing Then
                                Throw New Exception("WMA tags not found.")
                            Else
                                'Search for each requested tag
                                For Each MyTag As Tag In Tags
                                    Dim MatchFound As Asf.ContentDescriptor = Nothing

                                    For Each Field As Asf.ContentDescriptor In ASF
                                        Dim FieldName As String = Field.Name.Trim.ToUpper(InvariantCulture)
                                        If FieldName = MyTag.Name.ToUpper(InvariantCulture) Then
                                            MatchFound = Field
                                            Exit For
                                        ElseIf FieldName.Contains(MyTag.Name.ToUpper(InvariantCulture)) Then 'Could be a match, need to do an extra check...
                                            Dim TagSplit As String() = FieldName.Split("/"c)

                                            If TagSplit.Count > 1 AndAlso TagSplit(1).Trim.ToUpper(InvariantCulture) = MyTag.Name.ToUpper(InvariantCulture) Then
                                                MatchFound = Field
                                                Exit For
                                            End If
                                        End If
                                    Next

                                    If MatchFound IsNot Nothing Then 'If the value matches or wasn't requested, return true
                                        If MyTag.Value Is Nothing OrElse MyTag.Value = "" OrElse MyTag.Value.ToUpper(InvariantCulture) = ASF.GetDescriptorString(MatchFound.Name).Trim.ToUpper(InvariantCulture) Then
                                            Return New ReturnObject(True, "", True)
                                        End If
                                    End If
                                Next

                                'If none of the tags was found, return false
                                Return New ReturnObject(True, "", False)
                            End If
                        End Using

                        Return New ReturnObject(True, "", False)
                    End If
                Catch ex As Exception
                    Return New ReturnObject(False, ex.Message)
                End Try
            Else
                Return New ReturnObject(False, "Tags is nothing")
            End If

        End Function

    End Class

    Private Class MP3Codec

        Private Sub New()
            ' Unused...
        End Sub

        Public Shared Function MatchTag(FilePath As String, Tags As Tag()) As ReturnObject

            If Tags IsNot Nothing Then
                Try
                    If Tags.Length = 0 Then
                        'Tags weren't specified, so always match every file
                        Return New ReturnObject(True, "", True)
                    Else
                        'Find tags within MP3 file
                        Using MP3File As TagLib.File = TagLib.File.Create(FilePath)
                            Dim ID3 As Id3v2.Tag = CType(MP3File.GetTag(TagTypes.Id3v2, False), Id3v2.Tag)

                            If ID3 Is Nothing Then
                                Throw New Exception("MP3 tags not found.")
                            Else
                                Dim ID3Frames As List(Of Id3v2.Frame) = ID3.GetFrames.ToList
                                Dim ID3UserFrame As Id3v2.UserTextInformationFrame

                                'Search for each requested tag
                                For Each ID3Frame As Id3v2.Frame In ID3Frames
                                    Try 'Test if this is a user-defined ID3v2 tag - if not, skip to the next one
                                        ID3UserFrame = TryCast(ID3Frame, Id3v2.UserTextInformationFrame)

                                        If ID3UserFrame IsNot Nothing Then
                                            For Each MyTag As Tag In Tags
                                                If ID3UserFrame.Description.Trim.ToUpper(InvariantCulture) =
                                                    MyTag.Name.ToUpper(InvariantCulture) Then
                                                    'If the value matches or wasn't requested, return true
                                                    If MyTag.Value Is Nothing OrElse MyTag.Value = "" OrElse MyTag.Value.ToUpper(InvariantCulture) = ID3UserFrame.Text(0).Trim.ToUpper(InvariantCulture) Then
                                                        Return New ReturnObject(True, "", True)
                                                    End If
                                                End If
                                            Next
                                        End If
                                    Catch ex As Exception
                                        Continue For
                                    End Try
                                Next

                                'If none of the tags was found, return false
                                Return New ReturnObject(True, "", False)
                            End If
                        End Using

                        Return New ReturnObject(True, "", False)
                    End If
                Catch ex As Exception
                    Return New ReturnObject(False, ex.Message)
                End Try
            Else
                Return New ReturnObject(False, "Tags is nothing")
            End If

        End Function

    End Class

    Private Class AACCodec

        Private Sub New()
            ' Unused...
        End Sub

        Public Shared Function MatchTag(FilePath As String, Tags As Tag()) As ReturnObject

            Dim BOXTYPE_ILST As ReadOnlyByteVector = "ilst" 'List of tags
            Dim BOXTYPE_NAME As ReadOnlyByteVector = "name" 'Tag name
            Dim BOXTYPE_DATA As ReadOnlyByteVector = "data" 'Tag value

            If Tags IsNot Nothing Then
                Try
                    If Tags.Length = 0 Then
                        'Tags weren't specified, so always match every file
                        Return New ReturnObject(True, "", True)
                    Else
                        'Find tags within AAC file
                        Dim TagMatched As Boolean = False
                        'Grab user metadata box, which contains all of our tags
                        Using AAC_File As New Mpeg4TestFile(FilePath)
                            Dim UserDataBoxes As Mpeg4.IsoUserDataBox = AAC_File.UdtaBoxes(0)
                            Dim UserDataBox = DirectCast(UserDataBoxes.Children.First(), Mpeg4.IsoMetaBox)

                            'Search through each box until we find the "ilst" box
                            For a As Int32 = 0 To UserDataBox.Children.Count - 1
                                Try
                                    If UserDataBox.Children(a).BoxType = BOXTYPE_ILST Then
                                        'Search through child boxes of "ilst" box to find the relevant tags
                                        For Each UserData As Mpeg4.AppleAnnotationBox In CType(UserDataBox.Children(a), Mpeg4.AppleItemListBox).Children
                                            Dim TagFound As Tag = Nothing

                                            'If this AnnotationBox has children, look through them for tag data
                                            If UserData.Children.Count > 0 Then
                                                For Each TagBox In UserData.Children
                                                    If TagBox.BoxType = BOXTYPE_NAME Then

                                                        Debug.WriteLine(CType(TagBox, Mpeg4.AppleAdditionalInfoBox).Text.Replace(Convert.ToChar(0), "").Trim)

                                                        'This AppleAdditionalInfoBox contains the name of the tag, so look for it in our list of tag names
                                                        For Each MyTag As Tag In Tags
                                                            If CType(TagBox, Mpeg4.AppleAdditionalInfoBox).Text.Replace(Convert.ToChar(0), "").Trim.ToUpper(InvariantCulture) = MyTag.Name.ToUpper(InvariantCulture) Then
                                                                TagFound = MyTag
                                                                Exit For
                                                            End If
                                                        Next
                                                    ElseIf TagBox.BoxType = BOXTYPE_DATA Then
                                                        'This AppleAdditionalInfoBox contains the value of the tag, so if this tag was found in our tag list
                                                        'we need to check if the tag's value also matches (or that no specific value was requested)
                                                        If TagFound IsNot Nothing Then
                                                            If TagFound.Value Is Nothing OrElse TagFound.Value = "" OrElse CType(TagBox, Mpeg4.AppleDataBox).Text.Trim.ToUpper(InvariantCulture) = TagFound.Value.ToUpper(InvariantCulture) Then
                                                                TagMatched = True
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                Next
                                            End If

                                            'If we matched a tag, we can end our search now
                                            If TagMatched Then Exit For
                                        Next
                                    End If
                                Catch ex As Exception
                                    Return New ReturnObject(False, ex.Message)
                                End Try

                                'If we matched a tag, we can end our search now
                                If TagMatched Then Exit For
                            Next
                        End Using

                        If TagMatched Then
                            Return New ReturnObject(True, "", True)
                        Else
                            Return New ReturnObject(True, "", False)
                        End If
                    End If
                Catch ex As Exception
                    Return New ReturnObject(False, ex.Message)
                End Try
            Else
                Return New ReturnObject(False, "Tags is nothing")
            End If

        End Function

    End Class

End Class
