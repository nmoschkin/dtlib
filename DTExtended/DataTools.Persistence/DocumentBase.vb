Option Explicit On
Option Compare Binary

Imports System.IO
Imports System.Text
Imports System.Windows.Forms
Imports System.Drawing
Imports DataTools.Interop.Native

Imports DataTools.Persistence
Imports System.ComponentModel
Imports System.Drawing.Design
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Json
Imports System.IO.Compression

<TypeConverter(GetType(ExpandableObjectConverter)), DataContract, Serializable()>
Public Class ProgramFileTypeCollection
    Inherits ObjectModel.ObservableCollection(Of ProgramFileType)

    <DataMember, Browsable(True), Category("Appearance")>
    Public Property CombinedFiltersDescription As String

    <DataMember, Browsable(True), Category("Appearance")>
    Public Property DocumentTypeDescription As String

    Public Function GenerateFilter() As String
        Dim sa As String = CombinedFiltersDescription & " ("
        Dim sb As New StringBuilder
        Dim sc As String = ""
        Dim c As Integer = 0

        For Each ft In Me
            If c <> 0 Then
                sa &= ";"
                sc &= ";"
                sb.Append("|")
            End If

            sa &= "*" & ft.Extension
            sc &= "*" & ft.Extension

            sb.Append(ft.Description & " (" & "*" & ft.Extension & ")|*" & ft.Extension)
            c += 1
        Next

        If c < 2 Then sa = sc Else sa &= ")|" & sc

        GenerateFilter = sb.ToString & "|All Files (*.*)|*.*"
    End Function

End Class

<DataContract>
<TypeConverter(GetType(ExpandableObjectConverter))>
Public Class ProgramFileType

    Protected _Extension As String
    Protected _Description As String

    <DataMember, Browsable(True), Category("Design"), Description("Extension of file.")>
    Public Property Extension As String
        Get
            Return _Extension
        End Get
        Set(value As String)
            If value.IndexOf(".") = -1 Then
                value = "." & value.ToLower
            ElseIf value.IndexOf(".") <> 0 Then
                value = value.Substring(value.ToLower.IndexOf("."))
            Else
                value = value.ToLower
            End If

            _Extension = value
        End Set
    End Property

    <DataMember, Browsable(True), Category("Appearance"), Description("Description of file type.")>
    Public Property Description As String
        Get
            Return _Description
        End Get
        Set(value As String)
            value = TitleCase(value.Trim)
            If value.Length < 5 OrElse value.Substring(value.Length - 5) <> "Files" Then
                value &= " Files"
            End If

            If _Description <> value Then
                _Description = value
            End If
        End Set
    End Property

    <DataMember, Browsable(True), Category("Design"), Description("Program description")>
    Public Property Program As String

    <DataMember, Browsable(True), Category("Appearance"), Description("Default icon for file type,")>
    Public Property Icon As Image

    <DataMember, Browsable(True), Category("Design"), Description("General format of file.")>
    Public Property FileFormat As SerializationTypes = SerializationTypes.Binary

End Class

<DataContract>
Public MustInherit Class JsonWritableDocument

    ''' <summary>
    ''' Save the entire object graph to a JSON file.
    ''' </summary>
    ''' <param name="fileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function SaveJson(fileName As String, Optional promptOverwrite As Boolean = False, Optional ignoreParentProperties As Boolean = False) As Boolean
        CleanExtension(fileName)

        If Path.GetExtension(fileName).ToLower <> ".json" Then
            fileName &= ".json"
        End If

        If promptOverwrite AndAlso File.Exists(fileName) Then
            If MsgBox("Do you want to overwrite '" & Path.GetFileName(fileName) & "'?", vbYesNo) = MsgBoxResult.No Then
                Return False
            End If
        End If

        Dim fs As New FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None)

        SaveJson = internalSaveJson(fs, ignoreParentProperties)

    End Function

    Public Overridable Function SaveJsonClass(fileName As String, Optional promptOverwrite As Boolean = False) As Boolean

        If promptOverwrite AndAlso File.Exists(fileName) Then
            If MsgBox("Do you want to overwrite '" & Path.GetFileName(fileName) & "'?", vbYesNo) = MsgBoxResult.No Then
                Return False
            End If
        End If

        Try
            Dim s As String = WriteJavaScriptObjectGraphTemplate(Me.GetType, Me)
            File.WriteAllText(fileName, s)

        Catch ex As Exception

            Return False
        End Try

        Return True

    End Function

    Protected Overridable Function internalSaveJson(stream As Stream, Optional ignoreParent As Boolean = False) As Boolean

        Try

            Dim js As DataContractJsonSerializer
            Dim jset As New DataContractJsonSerializerSettings

            Dim kt() As System.Type = DataTools.Persistence.Helpers.GetKnownTypeList(Me.GetType())

            ReDim Preserve kt(kt.Count)

            jset.KnownTypes = kt

            jset.SerializeReadOnlyTypes = True
            jset.EmitTypeInformation = EmitTypeInformation.Never

            js = New DataContractJsonSerializer(Me.GetType, jset)

            js.WriteObject(stream, Me)

            stream.Flush()
            stream.Close()

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

End Class

<DataContract>
<TypeConverter(GetType(ExpandableObjectConverter))>
Public MustInherit Class DocumentBase
    Inherits JsonWritableDocument
    Implements INotifyPropertyChanged

    Protected _Compress As Boolean = True

    Protected _FileTypes As ProgramFileTypeCollection

    Protected _DocumentFormat As SerializationTypes = SerializationTypes.Binary

    Protected WithEvents _Serializer As ISerializerAdapter

    Protected _CreationDate As Date = Now

    Protected _LastModifyDate As Date = Now

    Protected _DisableNotify As Boolean = False

    Protected _LastDirectory As String

    Public Event DocumentLoaded(sender As Object, e As EventArgs)

    Public Event DocumentSaved(sender As Object, e As EventArgs)

    ''' <summary>
    ''' Gets or sets the last directory that was open.
    ''' </summary>
    ''' <returns></returns>
    Public Property LastDirectory As String
        Get
            Return _LastDirectory
        End Get
        Set(value As String)
            _LastDirectory = value
            OnPropertyChanged("LastDirectory")
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets an object that implements the <see cref="ISerializerAdapter"/> interface.
    ''' You can use these in lieu of the default binary and XML documents to save your files.
    ''' Calling <see cref="DocumentBase.ResetSerializer"/> clears this custom setting.  To reset your custom serializer,
    ''' call the <see cref="DocumentBase.ResetSerializer(Of T, ISerializerAdapter)(Boolean, SerializerAdapterFlags, String, Long)"/> function, instead.
    ''' </summary>
    ''' <returns></returns>
    Public Property Serializer As ISerializerAdapter
        Get
            Return _Serializer
        End Get
        Set(value As ISerializerAdapter)
            If _Serializer Is value Then Return

            If _Serializer IsNot Nothing Then
                RemoveHandler _Serializer.InstanceHelper, AddressOf InstanceHelper
            End If

            _Serializer = value

            If value IsNot Nothing Then AddHandler _Serializer.InstanceHelper, AddressOf InstanceHelper
            OnPropertyChanged("Serializer")
        End Set
    End Property


    ''' <summary>
    ''' Gets or sets a value indicating whether PropertyChanged events will be fired
    ''' when values are changed in the document.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property DisableNotify As Boolean
        Get
            Return _DisableNotify
        End Get
        Set(value As Boolean)
            If (value <> _DisableNotify) Then
                _DisableNotify = value
                If value = False Then
                    OnPropertyChanged("DisableNotify")
                End If
            End If
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets a string monikor for the unique 64-bit type code for this file type.
    ''' Only applies to binary file formats.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TypeCode As String
        Get
            If Me.DocumentFormat <> SerializationTypes.Binary Then Return Me.DocumentFormat.ToString

            If _Serializer Is Nothing Then ResetSerializer()
            Return ASCIIEncoding.ASCII.GetString(BitConverter.GetBytes(_Serializer.FileCode))
        End Get
        Set(value As String)
            If _Serializer Is Nothing Then ResetSerializer(False)

            If Me.DocumentFormat <> SerializationTypes.Binary Then

                If Not [Enum].TryParse(Of SerializationTypes)(value, Me.DocumentFormat) Then
                    Me.DocumentFormat = SerializationTypes.Binary
                Else
                    Return
                End If

            End If

            If value.Length > 8 Then
                value = value.Substring(0, 8)
            ElseIf value.Length < 8 Then
                value = value + New String("0", 8 - value.Length)
            End If

            Dim l As Long = BitConverter.ToInt64(ASCIIEncoding.ASCII.GetBytes(value), 0)

            If l = _Serializer.FileCode Then Return
            _Serializer.FileCode = l

            OnPropertyChanged("TypeCode")
            OnPropertyChanged("FileCode")
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the unique 64-bit type code for this file type.
    ''' Only applies to binary file formats.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FileCode As Long
        Get
            Return _Serializer.FileCode
        End Get
        Set(value As Long)

            If _Serializer Is Nothing Then ResetSerializer(False)
            If _Serializer.FileCode = value Then Return

            _Serializer.FileCode = value

            OnPropertyChanged("TypeCode")
            OnPropertyChanged("FileCode")
        End Set
    End Property


    ''' <summary>
    ''' Gets or sets a value indicating whether or not the document has been changed since last write.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Browsable(False)>
    Public Property Changed As Boolean

    ''' <summary>
    ''' Gets or sets a property indicating that user interaction will not be attempted by this document object.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DataMember, Browsable(True), Category("Design")>
    Public Property SuppressUI As Boolean = False

    Private _Guid As System.Guid = System.Guid.NewGuid

    ''' <summary>
    ''' Gets or sets the GUID of this document. A GUID is a unique identifier that is used for comprehensive identification
    ''' </summary>
    ''' <remarks></remarks>
    <DataMember, Browsable(True), Category("Design"),
    Editor(GetType(GuidTypeEditor), GetType(UITypeEditor)),
    Description("Gets or sets the GUID of this document. A GUID is a unique identifier that is used for comprehensive identification.")>
    Public Property Guid As System.Guid
        Get
            Return _Guid
        End Get
        Set(value As System.Guid)
            _Guid = value
            OnPropertyChanged("Guid")
        End Set
    End Property

    Private _Title As String = "New Document"

    ''' <summary>
    ''' Gets or set the title of the document.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DataMember, Browsable(True), Category("Design"), Description("Gets or set the title of the document.")>
    Public Property Title As String
        Get
            Return _Title
        End Get
        Set(value As String)
            _Title = value
            OnPropertyChanged("Title")
        End Set
    End Property

    Private _Author As String = "Document Author"

    ''' <summary>
    ''' Gets or set the author of the document.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DataMember, Browsable(True), Category("Design"), Description("Gets or set the author of the document.")>
    Public Property Author As String
        Get
            Return _Author
        End Get
        Set(value As String)
            _Author = value
            OnPropertyChanged("Author")
        End Set
    End Property

    Private _Filename As String

    ''' <summary>
    ''' Gets or set the filename of the document.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DataMember, Browsable(True), Category("Data"), Description("Gets or set the filename of the document.")>
    Public Property FileName As String
        Get
            Return _Filename
        End Get
        Set(value As String)
            _Filename = value
            OnPropertyChanged("FileName")
        End Set
    End Property

    ''' <summary>
    ''' Gets the creation date of the document.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DataMember, Browsable(True), Category("Design"), Description("Gets the creation date of the document.")>
    Public Property CreationDate As Date
        Get
            Return _CreationDate
        End Get
        Set(value As Date)
            _CreationDate = value
            OnPropertyChanged("CreationDate")
        End Set
    End Property

    ''' <summary>
    ''' Gets the last modification date of the document.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DataMember, Browsable(True), Category("Design"), Description("Gets the last modification date of the document.")>
    Public Property LastModifyDate As Date
        Get
            Return _LastModifyDate
        End Get
        Set(value As Date)
            _LastModifyDate = value
            OnPropertyChanged("LastModifyDate")
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the serialization format for the document.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Browsable(False), Category("Design"), Description("Gets or sets the serialization format for the document.")>
    Public Property DocumentFormat As SerializationTypes
        Get
            Return _DocumentFormat
        End Get
        Set(value As SerializationTypes)
            If value <> _DocumentFormat Then
                _DocumentFormat = value
                ResetSerializer(True)
                OnPropertyChanged("DocumentFormat")
            End If
        End Set
    End Property

    ''' <summary>
    ''' Gets the file type object associated with this document.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Browsable(True), Category("Data")>
    Public Property FileTypes As ProgramFileTypeCollection
        Get
            Return _FileTypes
        End Get
        Set(value As ProgramFileTypeCollection)
            _FileTypes = value
            OnPropertyChanged("FileTypes")
        End Set
    End Property

    Private WithEvents _Version As New VersionInfoObject

    ''' <summary>
    ''' Gets or sets the version of the document.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DataMember, Browsable(True), Category("Design"), Editor(GetType(VersionObjectTypeEditor), GetType(UITypeEditor))>
    Public Property Version As VersionInfoObject
        Get
            Return _Version
        End Get
        Set(value As VersionInfoObject)
            _Version = value
            OnPropertyChanged("Version")
        End Set
    End Property


    ''' <summary>
    ''' Gets or sets a value indicating that the data to be read or written is compressed with GZip.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Browsable(True), Category("Behavior")>
    Public Property Compress As Boolean
        Get
            Return _Compress
        End Get
        Set(value As Boolean)
            _Compress = value
            OnPropertyChanged("Compress")
        End Set
    End Property


    ''' <summary>
    ''' Initializes or reinitializes a custom serializer, optionally copying settings from the old serializer, if possible.
    ''' </summary>
    ''' <param name="copySettings">Whether or not to copy settings from the old instance of the serializer, if possible.</param>
    ''' <param name="serializationOptions">Specifies the flags to be used for serialization.</param>
    ''' <param name="typeCode">Specifies the typecode string to use as the file's unique 64-bit identifier.</param>
    ''' <param name="fileCode">Specifies the file's unique 64-bit identifier. If this parameter is set, then <paramref name="typeCode"/> is not used.</param>
    Public Overridable Sub ResetSerializer(Optional ByVal copySettings As Boolean = False, Optional ByVal serializationOptions As SerializerAdapterFlags = SerializerAdapterFlags.DataMember Or SerializerAdapterFlags.MatchPublic Or SerializerAdapterFlags.NoFields Or SerializerAdapterFlags.MatchReadProperty Or SerializerAdapterFlags.MatchWriteProperty, Optional typeCode As String = "", Optional fileCode As Long = 0)

        Dim sernew As ISerializerAdapter = _Serializer

        Dim os As SerializerAdapterFlags = serializationOptions

        If sernew IsNot Nothing AndAlso copySettings Then
            os = _Serializer.Options
        End If

        If _Serializer IsNot Nothing Then
            RemoveHandler _Serializer.InstanceHelper, AddressOf InstanceHelper
        End If

        Select Case _DocumentFormat

            Case SerializationTypes.Text, SerializationTypes.Xml
                _Serializer = New XmlSerializerAdapter()
                If fileCode Then Me.FileCode = fileCode Else Me.TypeCode = typeCode

            Case SerializationTypes.Binary

                _Serializer = New BinarySerializerAdapter()
                If fileCode Then Me.FileCode = fileCode Else Me.TypeCode = typeCode

            Case Else
                Throw New ArgumentException("DocumentFormat: There is no built-in support for that file type.")

        End Select

        _Serializer.Options = os
        _Serializer.KnownTypes = Helpers.GetKnownTypeList(Me.GetType)

        GC.Collect(2)
        AddHandler _Serializer.InstanceHelper, AddressOf InstanceHelper

    End Sub


    ''' <summary>
    ''' Initializes or reinitializes a custom serializer, optionally copying settings from the old serializer, if possible.
    ''' </summary>
    ''' <typeparam name="T">Type type of creatable <see cref="ISerializerAdapter"/> to use.</typeparam>
    ''' <typeparam name="ISerializerAdapter">An interface for serializing and deserializing object graphs.</typeparam>
    ''' <param name="copySettings">Whether or not to copy settings from the old instance of the serializer, if possible.</param>
    ''' <param name="serializationOptions">Specifies the flags to be used for serialization.</param>
    ''' <param name="typeCode">Specifies the typecode string to use as the file's unique 64-bit identifier.</param>
    ''' <param name="fileCode">Specifies the file's unique 64-bit identifier. If this parameter is set, then <paramref name="typeCode"/> is not used.</param>
    Public Overridable Sub ResetSerializer(Of T As New, ISerializerAdapter)(Optional ByVal copySettings As Boolean = False, Optional ByVal serializationOptions As SerializerAdapterFlags = SerializerAdapterFlags.DataMember Or SerializerAdapterFlags.MatchPublic Or SerializerAdapterFlags.NoFields Or SerializerAdapterFlags.MatchReadProperty Or SerializerAdapterFlags.MatchWriteProperty, Optional typeCode As String = "", Optional fileCode As Long = 0)

        Dim sernew As ISerializerAdapter = _Serializer
        Dim os As SerializerAdapterFlags = serializationOptions

        If sernew IsNot Nothing AndAlso copySettings Then
            os = _Serializer.Options
        End If

        If _Serializer IsNot Nothing Then
            RemoveHandler _Serializer.InstanceHelper, AddressOf InstanceHelper
        End If

        _Serializer = New T()
        If fileCode Then Me.FileCode = fileCode Else Me.TypeCode = typeCode

        _Serializer.Options = os
        _Serializer.KnownTypes = Helpers.GetKnownTypeList(Me.GetType)

        GC.Collect(2)
        AddHandler _Serializer.InstanceHelper, AddressOf InstanceHelper

    End Sub

    Private _dperf As Boolean = False

    ''' <summary>
    ''' To do any cleaning up before loading a document.
    ''' </summary>
    ''' <param name="fileName">The filename being loaded.</param>
    ''' <remarks></remarks>
    Protected Overridable Sub BeforeLoadDocument(fileName As String, ByRef cancel As Boolean)
        Return
    End Sub

    ''' <summary>
    ''' To do any cleaning up after loading a document.
    ''' </summary>
    ''' <param name="fileName">The file name that was loaded.</param>
    ''' <param name="loadSucceeded">True if the load succeeded.</param>
    ''' <remarks></remarks>
    Protected Overridable Sub AfterLoadDocument(fileName As String, loadSucceeded As Boolean)
        Return
    End Sub

    ''' <summary>
    ''' Attempts to load the document using the filename stored in the properties or by prompting the user.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function LoadDocument() As Boolean
        Dim c As Boolean = False

        BeforeLoadDocument(FileName, c)
        If c = True Then Return False

        Dim pFile As String = FileName

        If pFile Is Nothing OrElse Not File.Exists(pFile) Then
            If SuppressUI Then
                AfterLoadDocument(FileName, False)
                Return False
            End If

            Dim ofd As New System.Windows.Forms.OpenFileDialog
            ofd.Filter = FileTypes.GenerateFilter
            ofd.Title = FileTypes.DocumentTypeDescription

            ofd.CheckPathExists = True
            ofd.CheckFileExists = True
            ofd.InitialDirectory = LastDirectory
            If ofd.ShowDialog <> DialogResult.OK Then
                AfterLoadDocument(FileName, False)
                Return False
            End If

            Return LoadDocument(ofd.FileName, False)
        End If

        Return LoadDocument(_Filename, False)
    End Function

    ''' <summary>
    ''' Load the document from a file.
    ''' </summary>
    ''' <param name="fileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function LoadDocument(fileName As String, Optional beforeLoadAction As Boolean = True) As Boolean

        Dim fs As New FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read)

        Return LoadDocument(fs, beforeLoadAction)


    End Function

    ''' <summary>
    ''' Load the document from a stream.
    ''' </summary>
    ''' <param name="stream"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function LoadDocument(stream As Stream, Optional beforeLoadAction As Boolean = True) As Boolean

        If beforeLoadAction Then
            Dim c As Boolean = False

            BeforeLoadDocument(FileName, c)
            If c = True Then Return False
        End If

        If _Serializer Is Nothing Then ResetSerializer()

        Try
            _Serializer.KnownTypes = Helpers.GetKnownTypeList(Me.GetType)

            If _Compress Then
                Dim mm2 As New MemoryStream
                Dim gz As New Compression.GZipStream(stream, Compression.CompressionMode.Decompress, True)
                gz.CopyTo(mm2)

                mm2.Seek(0, SeekOrigin.Begin)

                stream.Close()
                stream = mm2
            End If

            If _Serializer.Decommit(stream, Me) = False Then
                MsgBox("Invalid file format", MsgBoxStyle.Critical)
                AfterLoadDocument(FileName, False)
            End If
        Catch ex As Exception
            MsgBox("Invalid file format", MsgBoxStyle.Critical)
            AfterLoadDocument(FileName, False)
        End Try

        AfterLoadDocument(FileName, True)
        Return True

    End Function

    ''' <summary>
    ''' To do any cleaning up before saving a document.
    ''' </summary>
    ''' <param name="fileName">The filename being saveed.</param>
    ''' <remarks></remarks>
    Protected Overridable Sub BeforeSaveDocument(fileName As String, ByRef cancel As Boolean)
        Return
    End Sub

    ''' <summary>
    ''' To do any cleaning up after saving a document.
    ''' </summary>
    ''' <param name="fileName">The file name that was saveed.</param>
    ''' <param name="saveSucceeded">True if the save succeeded.</param>
    ''' <remarks></remarks>
    Protected Overridable Sub AfterSaveDocument(fileName As String, saveSucceeded As Boolean)
        Return
    End Sub

    ''' <summary>
    ''' Save the document.
    ''' </summary>
    ''' <param name="promptOverWrite">Whether to prompt the user to overwrite.  If SuppressUI is set, this method will return false.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function SaveDocument(promptOverWrite As Boolean) As Boolean
        Return SaveDocument(_Filename, promptOverWrite)
    End Function

    ''' <summary>
    ''' Saves the document.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function SaveDocument() As Boolean
        Return SaveDocument(False)
    End Function

    ''' <summary>
    ''' Opens a save document prompt and saves the file.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function SaveDocumentAs() As Boolean
        Return SaveDocument("", False)
    End Function

    ''' <summary>
    ''' Initialize the _FileTypes ProgramFileTypesCollection class.
    ''' </summary>
    ''' <remarks></remarks>
    Protected MustOverride Sub InitializeFileTypes()

    ''' <summary>
    ''' Saves the document to the specified file.
    ''' </summary>
    ''' <param name="fileName">The filename to save.</param>
    ''' <param name="promptOverWrite">Whether to prompt the user to overwrite.  If SuppressUI is set, this method will return false.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function SaveDocument(fileName As String, Optional promptOverWrite As Boolean = True) As Boolean
        Dim pFile As String = fileName

        If String.IsNullOrEmpty(pFile) = False Then

            If (File.Exists(pFile) AndAlso (promptOverWrite)) Then
                If SuppressUI Then Return False

                Select Case MsgBox("'" & pFile & "' exists." & vbCrLf & "Would you like to replace it?", MsgBoxStyle.YesNoCancel)
                    Case MsgBoxResult.Yes
                        Exit Select

                    Case MsgBoxResult.No
                        pFile = Nothing

                    Case MsgBoxResult.Cancel
                        Return False

                End Select
            End If

        Else

            If SuppressUI Then Return False

            Dim sfd As New System.Windows.Forms.SaveFileDialog
            sfd.Filter = FileTypes.GenerateFilter
            sfd.Title = FileTypes.DocumentTypeDescription

            If _Filename IsNot Nothing Then
                sfd.FileName = Path.GetFileName(_Filename)
            End If

            sfd.InitialDirectory = LastDirectory
            sfd.CheckPathExists = True
            sfd.OverwritePrompt = True

            If sfd.ShowDialog <> DialogResult.OK Then Return False
            pFile = sfd.FileName

        End If

        Dim fs As New FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None)
        Return Me.SaveDocument(fs, True)

    End Function

    ''' <summary>
    ''' Save the document to a stream.
    ''' </summary>
    ''' <param name="stream">The stream to save the document to.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function SaveDocument(stream As Stream) As Boolean
        Return SaveDocument(stream, True)
    End Function

    ''' <summary>
    ''' Save the document to a stream.
    ''' </summary>
    ''' <param name="stream">The stream to save the document to.</param>
    ''' <param name="closeStream">Specifies whether or not to close the stream.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function SaveDocument(stream As Stream, closeStream As Boolean) As Boolean

        Dim c As Boolean
        BeforeSaveDocument(FileName, c)
        If c Then Return False

        LastModifyDate = Now
        If _Serializer Is Nothing Then ResetSerializer()

        _Serializer.KnownTypes = Helpers.GetKnownTypeList(Me.GetType)

        If _Compress Then
            Dim mm2 As New MemoryStream
            Dim gz As New GZipStream(stream, Compression.CompressionMode.Compress, True)

            c = _Serializer.Commit(Me, mm2, False)
            If c Then mm2.CopyTo(gz)
            If closeStream Then
                gz.Close()
                stream.Close()
                mm2.Close()
            End If
        Else
            c = _Serializer.Commit(Me, stream)
            If closeStream Then stream.Close()
        End If

        If c Then
            RaiseEvent DocumentSaved(Me, New EventArgs)
        End If

        AfterSaveDocument(FileName, c)

        Return c
    End Function

    Protected Overrides Function internalSaveJson(stream As Stream, Optional ignoreParent As Boolean = False) As Boolean
        Me.DisableNotify = True
        Return MyBase.internalSaveJson(stream, ignoreParent)
        Me.DisableNotify = False
    End Function

    Public Sub New()
        InitializeFileTypes()
        ResetSerializer()
    End Sub

    Public Sub New(fileName As String)
        InitializeFileTypes()
        ResetSerializer()
        If LoadDocument(fileName) = False Then
            Throw New FileNotFoundException("Document cannot be loaded.", fileName)
        End If
    End Sub

    Public Sub New(stream As Stream)
        InitializeFileTypes()
        ResetSerializer()
        If Not LoadDocument(stream) Then
            Throw New NativeException
        End If
    End Sub

    ''' <summary>
    ''' Called to deserialize a specific object.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Overridable Sub InstanceHelper(sender As Object, e As InstanceHelperEventArgs)

        If e.Instance IsNot Nothing Then Return

        If e.TypeName = "System.String" Then
            e.Instance = ""
        Else
            Try
                e.Instance = TypeDescriptor.CreateInstance(Nothing, Type.GetType(e.TypeName), Nothing, Nothing)
            Catch ex As Exception

            End Try
            If e.Instance Is Nothing Then e.Instance = FormatterServices.GetSafeUninitializedObject(Type.GetType(e.TypeName))
        End If
    End Sub

    Protected Overridable Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        If Not _DisableNotify Then RaiseEvent PropertyChanged(Me, e)
    End Sub

    Protected Overridable Sub OnPropertyChanged(e As String)
        If Not _DisableNotify Then RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(e))
    End Sub

    Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements INotifyPropertyChanged.PropertyChanged

    Private Sub _Version_PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Handles _Version.PropertyChanged
        '    If Not _DisableNotify Then OnPropertyChanged("Version")
    End Sub

End Class