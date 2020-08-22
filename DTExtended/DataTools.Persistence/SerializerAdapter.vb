'' Serializer Adapters (Version 2.0)
'' Main Component Class

'' Copyright (C) 2014-2015 Nathaniel Moschkin
'' All Rights Reserved

'' SerializerAdapter (AdaptableSerializer v2.0)
'' Copyright (C) 2014-2015 Nathaniel Moschkin.
'' All Rights Reserved.

'' Interface-based all-purpose object serialization and deserialization
'' with built in formatters for both binary and XML document interaction
'' and granular serialization control.

Option Explicit On
Option Compare Binary
Option Strict Off

Imports System
Imports System.IO
Imports System.Text

Imports System.ComponentModel
Imports System.ComponentModel.Design
Imports System.ComponentModel.Design.Serialization
Imports System.Drawing.Design
Imports System.Drawing
Imports System.Windows.Forms.Design
Imports System.Runtime.InteropServices
Imports System.Runtime.Serialization
Imports System.Reflection
Imports System.Threading
Imports System.IO.Compression
Imports System.Numerics
Imports System.Xml
Imports DataTools.Memory
Imports DataTools.Memory.Internal
Imports DataTools.Interop.Native

Namespace Persistence

#Region "InstanceHelperEventArgs"

    Public Class InstanceHelperEventArgs
        Inherits EventArgs

        Private _Type As System.Type
        Private _Guid As System.Guid

        Private _TypeName As String
        Private _Instance As Object

        Private _Cancel As Boolean
        Private _SubElementCount As Integer

        Public Property SubElementCount As Integer
            Get
                Return _SubElementCount
            End Get
            Friend Set(value As Integer)
                _SubElementCount = value
            End Set
        End Property

        Public Property Cancel() As Boolean
            Get
                Return _Cancel
            End Get
            Set(ByVal value As Boolean)
                _Cancel = value
            End Set
        End Property

        Public Property Instance() As Object
            Get
                Return _Instance
            End Get
            Set(ByVal value As Object)
                _Instance = value
            End Set
        End Property

        Public ReadOnly Property TypeName() As String
            Get
                Return _TypeName
            End Get
        End Property

        Public ReadOnly Property PropertyType() As System.Type
            Get
                Return _Type
            End Get
        End Property

        Public ReadOnly Property PropertyTypeGuid() As System.Guid
            Get
                Return _Guid
            End Get
        End Property

        Friend Sub New(t As System.Guid)
            On Error Resume Next

            _Guid = t
        End Sub

        Friend Sub New(t As System.Type)
            On Error Resume Next

            _Type = t
            _TypeName = t.FullName
            _Guid = t.GUID
        End Sub

        Friend Sub New(t As String)
            On Error Resume Next

            _TypeName = t
            _Type = Type.GetType(t)
        End Sub

    End Class

#End Region

#Region "SerializerAdapterFlags"


    Public Enum SerializationTypes As Short
        Binary = 0
        Text = 1
        Xml = 2
        Json = 3
        Custom = 99
    End Enum

    Public Enum BlockTypes As Short
        None = 0
        [Class] = 1
        [Array] = 2
        [Collection] = 3
        [String] = 4
        [Object] = 5

        [Binary] = &H80
        [ArrayBlob] = &H100
    End Enum

    Public Enum PrimitiveTypes As Short
        [String] = 0
        [Binary] = 1
        [Short] = 2
        [Integer] = 4
        [Long] = 8
        [Decimal] = 16
    End Enum

    ''' <summary>
    ''' Flags used to control the building of the serialization map.
    ''' </summary>
    ''' <remarks></remarks>
    <Flags>
    Public Enum SerializerAdapterFlags
        ''' <summary>
        ''' If this flag is set, only browsable properties are returned (in a set containing properties)
        ''' </summary>
        ''' <remarks></remarks>
        MatchBrowsable = 1

        ''' <summary>
        ''' If this flag is set, properties must be writable to be returned.
        ''' </summary>
        ''' <remarks></remarks>
        MatchWriteProperty = 2

        ''' <summary>
        ''' If this flag is set, properties must be readable to be returned.
        ''' </summary>
        ''' <remarks></remarks>
        MatchReadProperty = 4

        ''' <summary>
        ''' If this flag is set, public properties and fields are returned.  If this flag is unset, they are omitted.
        ''' </summary>
        ''' <remarks></remarks>
        MatchPublic = 8

        ''' <summary>
        ''' If this flag is set, private properties and fields are returned.  If this flag is unset, they are omitted.
        ''' </summary>
        ''' <remarks></remarks>
        MatchPrivate = 16

        ''' <summary>
        ''' If this flag is set, then only serializble classes and properties are serialized.
        ''' </summary>
        ''' <remarks></remarks>
        MatchSerializable = 32

        ''' <summary>
        ''' If this flag is set, attempts to match cannonical field names to property names using a variety of heuristics.
        ''' </summary>
        ''' <remarks></remarks>
        MatchPropertyToField = 64

        ''' <summary>
        ''' If this flag is set, fields are excluded from the match.
        ''' </summary>
        ''' <remarks></remarks>
        NoFields = 128

        ''' <summary>
        ''' If this flag is set, properties are excluded from the match.
        ''' </summary>
        ''' <remarks></remarks>
        NoProperties = 256

        ''' <summary>
        ''' If this flag is set, a class much have the DataContractAttribute in order to be serialized.
        ''' </summary>
        ''' <remarks></remarks>
        DataContract = 512

        ''' <summary>
        ''' If this flag is set, a property or field must have the DataMember attribute in order to be serialized.
        ''' </summary>
        ''' <remarks></remarks>
        DataMember = 1024

        ''' <summary>
        ''' If this flag is set, only matched properties/fields are returned and unmatched fields and properties are excluded.
        ''' This is used in conjuction with the MatchPropertyToField flag.
        ''' </summary>
        ''' <remarks></remarks>
        MatchesOnly = 2048

        ''' <summary>
        ''' If this flag is set, then properties must have the DesignerSerializationVisibilityAttribute set to
        ''' a value that is not 'Hidden.'
        ''' </summary>
        ''' <remarks></remarks>
        MatchDesignerVisibility = 4096

        ''' <summary>
        ''' If this flag is set, properties can match either Browsable or Designer-visible to be accepted.
        ''' </summary>
        ''' <remarks></remarks>
        MatchBrowsableOrVisible = 8192

        ''' <summary>
        ''' If a field and a property are matched, use the property name instead of the field name.
        ''' </summary>
        ''' <remarks></remarks>
        MatchUsePropName = 16384

    End Enum

    ''' <summary>
    ''' Visibility type constant for visibility-oriented attribute types
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum VisibilityType
        ''' <summary>
        ''' Attribute has a flag that indicates invisibility
        ''' </summary>
        ''' <remarks></remarks>
        Invisible = -1

        ''' <summary>
        ''' The attribute is not present.
        ''' </summary>
        ''' <remarks></remarks>
        Absent = 0

        ''' <summary>
        ''' Attribute has a flag that indicates visibility.
        ''' </summary>
        ''' <remarks></remarks>
        Visible = 1
    End Enum

#End Region

#Region "DateTime Patterns"

    Public Enum DateTimePatterns As Short
        ''' <summary>
        ''' Format to a resolution of ten millionths of a second in the pattern of:
        ''' yyyyMMddTHH:mm:ss.fffffffzz
        ''' </summary>
        ''' <remarks></remarks>
        WindowsVerbose = AscW("o")

        ''' <summary>
        ''' Format time according to the RFC1123 standard in the pattern of:
        ''' ddd, dd MMM yyyy HH:mm:ss GMT
        ''' </summary>
        ''' <remarks></remarks>
        RFC1123 = AscW("r")

        ''' <summary>
        ''' Format time according to the sortable DateTime in the pattern of:
        ''' yyyy-MM-ddTHH:mm:ss
        ''' </summary>
        ''' <remarks></remarks>
        SortableDateTimePattern = AscW("s")

        ''' <summary>
        ''' Format time according to the universal sortable DateTime in the pattern of:
        ''' yyyy-MM-dd HH:mm:ssZ
        ''' </summary>
        ''' <remarks></remarks>
        UniversalSortableDateTimePattern = AscW("u")
    End Enum

#End Region

#Region "Type Cache"

    Public Structure TypeSerializableMembers
        Public Prop As PropertyInfo
        Public Field As FieldInfo
        Public Description As String

        Public ReadOnly Property Name As String
            Get
                If Prop IsNot Nothing Then Return Prop.Name
                If Field IsNot Nothing Then Return Field.Name
                Return Description
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return Name
        End Function

    End Structure

    Public Structure TypeMember
        Public Type As System.Type
        Public Links() As TypeSerializableMembers

        Public Sub Add(tml As TypeSerializableMembers)
            If Links Is Nothing Then
                Links = {tml}
                Return
            End If

            If IndexOf(tml.Name) = -1 Then Return

            ReDim Preserve Links(Links.Count)
            Links(Links.Count - 1) = tml
        End Sub

        Public Function Remove(Name As String) As Boolean
            Dim x As Integer = IndexOf(Name)

            If Links Is Nothing Then Return False
            If x = -1 Then Return False

            Dim i As Integer, _
                c As Integer = Links.Count - 1

            Dim ln() As TypeSerializableMembers = Nothing
            Dim e As Integer = 0

            For i = 0 To c
                If x <> i Then
                    ReDim Preserve ln(e)
                    ln(e) = Links(i)
                    e += 1
                End If
            Next

            Links = ln
            Return True
        End Function

        Public Function IndexOf(Name As String) As Integer
            If Links Is Nothing OrElse Links.Count = 0 Then Return -1
            Dim i As Integer = 0

            For Each t As TypeSerializableMembers In Links
                If t.Prop IsNot Nothing Then
                    If t.Prop.Name = Name Then Return i
                End If

                If t.Field IsNot Nothing Then
                    If t.Field.Name = Name Then Return i
                End If

                i += 1
            Next

            Return -1
        End Function

    End Structure

    Public Structure TypeCache
        Public Members() As TypeMember

        Public Sub Add(t As System.Type, m() As MemberInfo)

            Dim mbr As New TypeMember
            mbr.Type = t

            For Each mx In m
                Dim tml As New TypeSerializableMembers

                If TypeOf mx Is FieldInfo Then
                    tml.Field = mx
                    tml.Description = GetDescription(tml.Field)
                ElseIf TypeOf mx Is PropertyInfo Then
                    tml.Prop = mx
                    tml.Description = GetDescription(tml.Prop)
                End If

                mbr.Add(tml)
            Next

        End Sub

        Public Sub Add(mbr As TypeMember)
            If Members Is Nothing Then
                Members = {mbr}
                Return
            End If

            If IndexOf(mbr.Type) = -1 Then Return

            ReDim Preserve Members(Members.Count)
            Members(Members.Count - 1) = mbr
        End Sub

        Public Function Remove(t As Type) As Boolean
            Dim x As Integer = IndexOf(t)

            If Members Is Nothing Then Return False
            If x = -1 Then Return False

            Dim i As Integer, _
                c As Integer = Members.Count - 1

            Dim ln() As TypeMember = Nothing
            Dim e As Integer = 0

            For i = 0 To c
                If x <> i Then
                    ReDim Preserve ln(e)
                    ln(e) = Members(i)
                    e += 1
                End If
            Next

            Members = ln
            Return True
        End Function

        Public Function IndexOf(t As Type) As Integer
            If Members Is Nothing OrElse Members.Count = 0 Then Return -1
            Dim i As Integer = 0

            For Each tp As TypeMember In Members
                If tp.Type.AssemblyQualifiedName = t.AssemblyQualifiedName Then Return i
                i += 1
            Next

            Return -1
        End Function

        Public Function Retrieve(t As Type) As TypeMember
            Dim i As Integer = IndexOf(t)
            If i = -1 Then Return Nothing
            Return Members(i)
        End Function

    End Structure

#End Region

#Region "ISelfSerializingObject"

    ''' <summary>
    ''' Provides an interface for a complex object that can more efficiently
    ''' serialize and deserialize itself instead of having the infrastructure attempt to figure it out.
    ''' </summary>
    Public Interface ISelfSerializingObject

        Function Serialize() As Byte()

        Sub Deserialize(data As Byte())

    End Interface

#End Region

#Region "ISerializerAdapter"

    ''' <summary>
    ''' Public interface for serializers.
    ''' </summary>
    ''' <remarks></remarks>
    Public Interface ISerializerAdapter
        ''' <summary>
        ''' Gets the serialization type implemented by the class.  Text = 0, Binary = 1, Xml = 2
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        ReadOnly Property Type As DataTools.Persistence.SerializationTypes

        ''' <summary>
        ''' Gets or sets the option flags used to determine the seriailzation pattern.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Property Options As SerializerAdapterFlags

        ''' <summary>
        ''' A list of types that are known to the serializer.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Property KnownTypes As System.Type()

        ''' <summary>
        ''' Commit an object tree to a file through serialization.
        ''' </summary>
        ''' <param name="obj">The root of the object tree to commit.</param>
        ''' <param name="fileName">The full name of the to write.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function Commit(obj As Object, fileName As String) As Boolean

        ''' <summary>
        ''' Commit an object tree to a stream through serialization.
        ''' </summary>
        ''' <param name="obj">The root of the object tree to commit.</param>
        ''' <param name="stream">The stream to which to commit the object.</param>
        ''' <param name="closeStream">Specifies whether to close the stream when the operation is complete.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function Commit(obj As Object, stream As System.IO.Stream, Optional ByVal closeStream As Boolean = True) As Boolean

        ''' <summary>
        ''' Decommit an object tree from a file through deserialization.
        ''' </summary>
        ''' <param name="fileName">The full name of the file to read.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function Decommit(fileName As String) As Object

        ''' <summary>
        ''' Decommit an object tree from a stream.
        ''' </summary>
        ''' <param name="stream">The stream from which to read.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function Decommit(stream As System.IO.Stream) As Object

        ''' <summary>
        ''' Decommit serialized data to a root object that may or not be preinstantiated.
        ''' </summary>
        ''' <param name="fileName">The full name of the file to read.</param>
        ''' <param name="rootObject">The object to receive the decommited data.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function Decommit(fileName As String, ByRef rootObject As Object) As Boolean

        ''' <summary>
        ''' Decommit serialized data to a root object that may or not be preinstantiated.
        ''' </summary>
        ''' <param name="stream">The stream from which to read.</param>
        ''' <param name="rootObject">The object to receive the decommited data.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function Decommit(stream As System.IO.Stream, ByRef rootObject As Object) As Boolean

        ''' <summary>
        ''' Reset the serializer to a just-create state, clearing all buffers and work items.
        ''' </summary>
        ''' <remarks></remarks>
        Sub Reset()

        ''' <summary>
        ''' A number that uniquely identifies the file type.
        ''' </summary>
        ''' <returns></returns>
        Property FileCode As Long

        Event InstanceHelper(sender As Object, e As InstanceHelperEventArgs)

    End Interface

#End Region

#Region "ISA Services"

    <HideModuleName>
    Public Module ISAServices

        ''' <summary>
        ''' Copies the settings from one serializer to another.  Usually used to change formats
        ''' while preserving logic.
        ''' </summary>
        ''' <param name="object1"></param>
        ''' <param name="object2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ISACopySettings(object1 As ISerializerAdapter, object2 As ISerializerAdapter) As Boolean

            Dim adapter As New BinarySerializerAdapter
            Dim stream As New MemoryStream

            ISACopySettings = adapter.Commit(object1, stream, False)

            If ISACopySettings Then
                ISACopySettings = adapter.Decommit(stream, object2)
            Else
                stream.Close()
            End If

        End Function

        Public Class CachedObject
            Public Property Type As System.Type

            Public Property Flags As SerializerAdapterFlags

            Public Property Members As MemberInfo()
        End Class

        Public NotInheritable Class Cache

            <ThreadStatic>
            Private Shared _cache As New List(Of CachedObject)

            Public Shared Sub Reset()
                _cache = New List(Of CachedObject)
            End Sub

            Public Shared Function Find(type As System.Type) As CachedObject
                If _cache Is Nothing Then Reset()

                For Each c In _cache
                    If c.Type = type Then Return c
                Next

                Return Nothing
            End Function

            Public Shared Function FindMembers(type As System.Type) As MemberInfo()
                If _cache Is Nothing Then Reset()

                For Each c In _cache
                    If c.Type = type Then Return c.Members
                Next

                Return Nothing
            End Function

            Public Shared Function FindMembers(type As System.Type, flags As SerializerAdapterFlags) As MemberInfo()
                If _cache Is Nothing Then Reset()

                For Each c In _cache
                    If c.Type = type AndAlso c.Flags = flags Then Return c.Members
                Next

                Return Nothing
            End Function

            Public ReadOnly Property Cache As List(Of CachedObject)
                Get
                    Return _cache
                End Get
            End Property

            Public Shared Sub Add(co As CachedObject)
                If _cache Is Nothing Then Reset()
                _cache.Add(co)
            End Sub

            Public Shared Sub Add(type As System.Type, flags As SerializerAdapterFlags, members As MemberInfo())
                If _cache Is Nothing Then Reset()
                _cache.Add(New CachedObject With {.Flags = flags, .Type = type, .Members = members})
            End Sub

            Public Shared Sub Clear()
                If _cache Is Nothing Then Reset()
                _cache.Clear()
            End Sub

        End Class


        ''' <summary>
        ''' Qualification Flags for the <see cref="FindKnownTypeLike(Type(), String, FindKnownTypeFlags)" /> function.
        ''' </summary>
        <Flags>
        Public Enum FindKnownTypeFlags

            ''' <summary>
            ''' The assembly name will be compared.
            ''' </summary>
            UseAssembly = 1

            ''' <summary>
            ''' The assembly version will be compared.
            ''' </summary>
            UseVersion = 2

            ''' <summary>
            ''' The assembly culture will be compared.
            ''' </summary>
            UseCulture = 4

            ''' <summary>
            ''' The assembly's public key will be compared.
            ''' </summary>
            UsePublicKey = 8
        End Enum

        ''' <summary>
        ''' Finds a known type from a given list that compares positively to the desired type according to the specified flags.
        ''' </summary>
        ''' <param name="kt">List of known types.</param>
        ''' <param name="type">Either the simple name or the Assembly Qualified name of the desired type.</param>
        ''' <param name="flags">The criteria for a positive match.</param>
        ''' <returns>Either a positively-matched type or Nothing.</returns>
        Public Function FindKnownTypeLike(kt() As System.Type, type As String, Optional flags As FindKnownTypeFlags = 1) As System.Type
            Dim s = BatchParse(type, ", ")
            Dim t() As String

            For Each k In kt
                t = BatchParse(k.AssemblyQualifiedName, ", ")

                If t(0) <> s(0) Then Continue For

                If s.Length < 4 Then Continue For
                If flags And FindKnownTypeFlags.UseAssembly Then
                    If t(1) <> s(1) Then Continue For
                End If

                If flags And FindKnownTypeFlags.UseVersion Then
                    If t(2) <> s(2) Then Continue For
                End If

                If flags And FindKnownTypeFlags.UseCulture Then
                    If t(3) <> s(3) Then Continue For
                End If

                If flags And FindKnownTypeFlags.UsePublicKey Then
                    If t(4) <> s(4) Then Continue For
                End If

                Return k
            Next

            Return Nothing
        End Function


        ''' <summary>
        ''' Create a map of all serializable members according to the specified flags.
        ''' </summary>
        ''' <param name="flags"></param>
        ''' <param name="compositeType"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function MakeSerializeMap(flags As SerializerAdapterFlags, compositeType As System.Type, ByRef matchedProps() As PropertyInfo) As MemberInfo()

            Dim cac = Cache.FindMembers(compositeType, flags)
            If cac IsNot Nothing Then Return cac

            Blob.ParseStrings = False

            '' check if the type, itself, can be serialized.

            '' check for data contract.
            If flags And SerializerAdapterFlags.DataContract Then
                If Not HasAttribute(compositeType, GetType(DataContractAttribute)) Then Return Nothing
            End If

            '' check for serializable attribute.
            If flags And SerializerAdapterFlags.MatchSerializable Then
                If Not HasAttribute(compositeType, GetType(SerializableAttribute)) Then Return Nothing
            End If

            '' retrieve all members from the type.
            Dim mi() As MemberInfo

            Dim bf As BindingFlags = BindingFlags.Instance

            If flags And SerializerAdapterFlags.MatchPrivate Then
                bf = bf Or BindingFlags.NonPublic
            End If

            If flags And SerializerAdapterFlags.MatchPublic Then
                bf = bf Or BindingFlags.Public
            End If

            mi = compositeType.GetMembers(bf)

            Dim mo() As MemberInfo = Nothing, _
                c As Integer = 0

            '' no members, not a composite type!
            If mi Is Nothing OrElse mi.Count = 0 Then Return Nothing

            '' perform a check on every field and property to get the first-pass list
            '' of serializable members.
            For Each mbr In mi

                If (mbr.MemberType = MemberTypes.Field AndAlso (0 = (flags And SerializerAdapterFlags.NoFields))) OrElse (mbr.MemberType = MemberTypes.Property) Then

                    If CanSerialize(flags, mbr) Then
                        ReDim Preserve mo(c)
                        mo(c) = mbr
                        c += 1
                    End If
                End If

            Next

            '' check for matched fields/properties.  If a match is found, only the fields
            '' will be serialized.  The properties will be discarded unless the MatchUseProperties
            '' is set to true, in which case the fields will be discarded.

            '' if we're matching properties to fields, we need to make sure that
            '' NoFields and NoProperties are unset, because we would have nothing to
            '' match in this list from the first pass, if they were.
            If (flags And SerializerAdapterFlags.MatchPropertyToField) AndAlso ((flags And SerializerAdapterFlags.NoFields) = 0) AndAlso ((flags And SerializerAdapterFlags.NoProperties) = 0) Then
                Dim fi() As FieldInfo = Nothing, pi() As PropertyInfo = Nothing, nomFi() As FieldInfo = Nothing, nomPi() As PropertyInfo = Nothing

                '' fi - matched fields
                '' pi - matched properties - passed to caller
                '' nomFi - unmatched fields
                '' nomPi - unmateched properties
                MatchFieldsProps(mo, fi, pi, nomFi, nomPi)
                mo = Nothing

                matchedProps = pi

                '' check if we have a matched field list.
                c = 0
                If fi IsNot Nothing Then
                    For Each fe In fi
                        ReDim Preserve mo(c)
                        mo(c) = fe
                        c += 1
                    Next
                End If

                '' if we only want matches, we don't want to add
                '' other fields and properties to the mix.
                If (flags And SerializerAdapterFlags.MatchesOnly) = 0 Then

                    '' add unmatched fields
                    If nomFi IsNot Nothing Then
                        For Each fe In nomFi
                            ReDim Preserve mo(c)
                            mo(c) = fe
                            c += 1
                        Next
                    End If

                    '' add unmatched properties.
                    If nomPi IsNot Nothing Then
                        For Each pe In nomPi
                            ReDim Preserve mo(c)
                            mo(c) = pe
                            c += 1
                        Next
                    End If

                End If

            End If

            Cache.Add(compositeType, flags, mo)
            MakeSerializeMap = mo

        End Function

        ''' <summary>
        ''' Determines, using the flags set, whether a property or field should be serialized.
        ''' </summary>
        ''' <param name="flags"></param>
        ''' <param name="info"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function CanSerialize(flags As SerializerAdapterFlags, info As MemberInfo) As Boolean
            Dim f As Boolean = True

            '' check if it's a field.
            If flags And SerializerAdapterFlags.NoFields Then
                If info.MemberType = MemberTypes.Field Then Return False
            End If

            '' check if it's a property.
            If flags And SerializerAdapterFlags.NoProperties Then
                If info.MemberType = MemberTypes.Property Then Return False
            End If

            '' check if it's a datamember.
            If flags And SerializerAdapterFlags.DataMember Then
                If Not HasAttribute(info, GetType(DataMemberAttribute)) Then Return Nothing
            End If

            '' check if it's private.
            If (flags And SerializerAdapterFlags.MatchPrivate) = 0 Then
                If IsPrivate(info) Then Return False
            End If

            '' check if it's public.
            If (flags And SerializerAdapterFlags.MatchPublic) = 0 Then
                If IsPublic(info) Then Return False
            End If

            If info.MemberType = MemberTypes.Property Then
                '' do some checks that apply only to properties.
                '' these flags are all opt-in, meaning a check
                '' will not occur unless the flag is set.

                '' check for read capability.
                If flags And SerializerAdapterFlags.MatchReadProperty Then
                    If CType(info, PropertyInfo).GetMethod Is Nothing Then
                        Return False
                    End If
                End If

                '' check for write capability.
                If flags And SerializerAdapterFlags.MatchWriteProperty Then
                    If CType(info, PropertyInfo).SetMethod Is Nothing Then
                        Return False
                    End If
                End If

                '' check for browsable or designer visibility.
                If flags And SerializerAdapterFlags.MatchBrowsableOrVisible Then
                    If Browsable(info) = VisibilityType.Visible Or DesignerSerialization(info) = VisibilityType.Visible Then
                        '' we can return, here.  the last two checks are obviated by assertion of this flag.
                        CanSerialize = True
                        Exit Function
                    End If
                End If

                '' check for Browsable visibility.
                If flags And SerializerAdapterFlags.MatchBrowsable Then
                    If Browsable(info) <> VisibilityType.Visible Then Return False
                End If

                '' check for designer visibility.
                If flags And SerializerAdapterFlags.MatchDesignerVisibility Then
                    If DesignerSerialization(info) <> VisibilityType.Visible Then Return False
                End If

            End If

            '' all checks pass, return true.
            CanSerialize = True

        End Function

        Public Function IsNameMatch(fieldName As String, propertyName As String) As Boolean
            Dim s As String

            '' Attempt to match a field with a property.
            '' Common Visual Basic schemes

            '' The automatic scheme.
            s = "_" & propertyName
            If fieldName = s Then
                Return True
            End If

            '' The old-style scheme carried over from C.
            s = "m_" & propertyName
            If fieldName = s Then
                Return True
            End If

            '' Another common scheme.
            s = "m" & propertyName
            If fieldName = s Then
                Return True
            End If

            '' The C# way of doing it is to use mixed capitals and lowercase.
            s = propertyName.ToLower
            If fieldName.ToLower = s Then
                Return True
            End If

            IsNameMatch = False
        End Function

        Public Function MatchPropertyToField(pname As String, fi() As FieldInfo) As FieldInfo
            Dim s As String = ""
            Dim fe As FieldInfo = Nothing

            For Each fe In fi
                If IsNameMatch(fe.Name, pname) Then Exit For
                fe = Nothing
            Next

            MatchPropertyToField = fe
        End Function

        Public Function MatchPropertyToField(pe As PropertyInfo, fi() As FieldInfo) As FieldInfo
            MatchPropertyToField = MatchPropertyToField(pe.Name, fi)
        End Function

        Public Function MatchFieldToProperty(fname As String, pi() As PropertyInfo) As PropertyInfo
            Dim s As String = ""
            Dim pe As PropertyInfo = Nothing

            For Each pe In pi
                If IsNameMatch(fname, pe.Name) Then Exit For
                pe = Nothing
            Next

            MatchFieldToProperty = pe
        End Function

        Public Function MatchFieldToProperty(pe As PropertyInfo, pi() As PropertyInfo) As PropertyInfo
            MatchFieldToProperty = MatchFieldToProperty(pe.Name, pi)
        End Function

        Private Sub MatchFieldsProps(mi() As MemberInfo, ByRef fi() As FieldInfo, ByRef pi() As PropertyInfo, ByRef nomFi() As FieldInfo, ByRef nomPi() As PropertyInfo)

            Dim c As Integer = 0
            Dim d As Integer = 0

            '' a switch to indicate a find.
            Dim m As Boolean = False

            Dim s As String = ""
            Dim fe As FieldInfo
            Dim pe As PropertyInfo = Nothing

            '' matched fields
            Dim fio() As FieldInfo = Nothing

            '' matched properties
            Dim pio() As PropertyInfo = Nothing

            '' unmatched fields
            Dim nfio() As FieldInfo = Nothing

            '' unmatched properties
            Dim npio() As PropertyInfo = Nothing

            Dim gf() As FieldInfo = GetFields(mi)
            Dim gp() As PropertyInfo = GetProperties(mi)

            '' perform the double loop initial match:
            For Each fe In gf

                pe = MatchFieldToProperty(fe.Name, gp)

                '' we have found the match, save both the property and field for return.
                If pe IsNot Nothing Then
                    ReDim Preserve pio(c)
                    ReDim Preserve fio(c)

                    fio(c) = fe
                    pio(c) = pe

                    c += 1
                Else
                    '' no match, here, we'll take this field and put it in the unmatched list.
                    ReDim Preserve nfio(d)
                    nfio(d) = fe
                    d += 1
                End If
            Next

            '' now find the unmatched properties.
            d = 0
            For Each pe In gp
                m = False

                If pio IsNot Nothing Then
                    For Each pe2 In pio
                        If pe2 Is pe Then
                            m = True
                            Exit For
                        End If
                    Next
                End If

                If Not m Then
                    ReDim Preserve npio(d)
                    npio(d) = pe
                    d += 1
                End If
            Next

            '' assign the work variables to the output variables.
            pi = pio
            fi = fio
            nomFi = nfio
            nomPi = npio

        End Sub

        ''' <summary>
        ''' Retrieve all properties from a MemberInfo array.
        ''' </summary>
        ''' <param name="mi"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetProperties(mi() As MemberInfo) As PropertyInfo()
            Dim pi() As PropertyInfo, _
                c As Integer = 0

            ReDim pi(mi.Length - 1)

            For Each mbr In mi
                If mbr.MemberType = MemberTypes.Property Then
                    pi(c) = mbr
                    c += 1
                End If
            Next

            ReDim Preserve pi(c - 1)
            GetProperties = pi
        End Function

        ''' <summary>
        ''' Retrieve all fields from a MemberInfo array.
        ''' </summary>
        ''' <param name="mi"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetFields(mi() As MemberInfo) As FieldInfo()
            Dim fi() As FieldInfo, _
                c As Integer = 0

            ReDim fi(mi.Length - 1)

            For Each mbr In mi
                If mbr.MemberType = MemberTypes.Field Then
                    fi(c) = mbr
                    c += 1
                End If
            Next

            ReDim Preserve fi(c - 1)
            GetFields = fi
        End Function

        ''' <summary>
        ''' Tests if a property or field is public.
        ''' </summary>
        ''' <param name="inf"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function IsPublic(inf As Object) As Boolean
            Dim isPub As Boolean = False

            Select Case inf.GetType

                ''unbox
                Case GetType(FieldInfo)
                    ''unbox
                    isPub = CType(inf, FieldInfo).IsPublic

                    ''unbox
                Case GetType(PropertyInfo)

                    ''unbox
                    If CType(inf, PropertyInfo).GetMethod IsNot Nothing Then
                        ''unbox
                        isPub = CType(inf, PropertyInfo).GetMethod.IsPublic
                    Else
                        ''unbox
                        isPub = CType(inf, PropertyInfo).SetMethod.IsPublic
                    End If

                Case Else
                    Return False

            End Select

            IsPublic = isPub
        End Function

        ''' <summary>
        ''' Tests if a property or field is private.
        ''' </summary>
        ''' <param name="inf"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function IsPrivate(inf As Object) As Boolean
            Dim isPriv As Boolean = False

            Select Case inf.GetType

                ''unbox
                Case GetType(FieldInfo)
                    ''unbox
                    isPriv = CType(inf, FieldInfo).IsPrivate

                    ''unbox
                Case GetType(PropertyInfo)

                    ''unbox
                    If CType(inf, PropertyInfo).GetMethod IsNot Nothing Then
                        ''unbox
                        isPriv = CType(inf, PropertyInfo).GetMethod.IsPrivate
                    Else
                        ''unbox
                        isPriv = CType(inf, PropertyInfo).SetMethod.IsPrivate
                    End If

                Case Else
                    Return False

            End Select

            IsPrivate = isPriv
        End Function


        Public Function GetDescription(mbr As MemberInfo) As String
            Dim fi As FieldInfo
            Dim pi As PropertyInfo

            Dim o As Object


            If TypeOf mbr Is FieldInfo Then
                fi = mbr
                o = fi.GetCustomAttribute(GetType(DescriptionAttribute))
                If o IsNot Nothing Then Return CType(o, DescriptionAttribute).Description()

            ElseIf TypeOf mbr Is PropertyInfo Then
                pi = mbr
                o = pi.GetCustomAttribute(GetType(DescriptionAttribute))
                If o IsNot Nothing Then Return CType(o, DescriptionAttribute).Description()
            End If

            Return Nothing
        End Function

        ''' <summary>
        ''' Returns true if the specified object has the specified custom attribute.
        ''' </summary>
        ''' <param name="inf"></param>
        ''' <param name="attr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function HasAttribute(inf As Object, attr As System.Type) As Boolean
            Dim ift As IEnumerable(Of CustomAttributeData) = Nothing

            If TypeOf inf Is System.Type Then

                If CType(inf, System.Type).GetCustomAttribute(attr) IsNot Nothing Then
                    Return True
                Else
                    Return False
                End If

            ElseIf TypeOf inf Is PropertyDescriptor Then
                For Each cattr In CType(inf, PropertyDescriptor).Attributes
                    If cattr.GetType = attr Then Return True
                Next

                Return False
            Else
                '' Custom attributes are new to .NET 4.5
                Try
                    Dim mbt As MemberInfo = TryCast(inf, MemberInfo)

                    If mbt Is Nothing Then
                        '' not a MemberInfo object, just get the GetType.CustomAttributes
                        ift = inf.GetType.CustomAttributes
                    Else
                        Select Case mbt.MemberType

                            Case MemberTypes.Field
                                ''unbox
                                ift = CType(inf, FieldInfo).CustomAttributes

                            Case MemberTypes.Property
                                ''unbox
                                ift = CType(inf, PropertyInfo).CustomAttributes

                            Case Else

                        End Select

                    End If

                Catch ex As Exception
                    '' not a MemberInfo object, just get the GetType.CustomAttributes
                    ift = inf.GetType.CustomAttributes
                End Try

            End If

            '' still nothing?  return false.
            If ift Is Nothing Then Return False

            For Each cattr In ift
                '' we have found an attribute that matches
                If cattr.AttributeType = attr Then Return True
            Next

            '' nothing, return false.
            Return False
        End Function

        ''' <summary>
        ''' Gets a value indicating designer serialization visibility. 1 is visible, -1 is not visible, and 0 is not present.
        ''' </summary>
        ''' <param name="inf"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function DesignerSerialization(inf As Object) As VisibilityType
            DesignerSerialization = VisibilityType.Absent

            Dim ift As IEnumerable(Of CustomAttributeData) = Nothing
            Try
                '' get a MemberInfo objects attributes.
                Dim mbt As MemberInfo = TryCast(inf, MemberInfo)
                Select Case mbt.MemberType

                    Case MemberTypes.Field
                        ''unbox
                        ift = CType(inf, FieldInfo).CustomAttributes

                    Case MemberTypes.Property
                        ''unbox
                        ift = CType(inf, PropertyInfo).CustomAttributes

                    Case Else

                End Select

            Catch ex As Exception
                '' get the other kind of object's attributes
                ift = inf.GetType.CustomAttributes
            End Try

            If ift Is Nothing Then Return False

            For Each cattr In ift
                If cattr.AttributeType = GetType(DesignerSerializationVisibilityAttribute) Then
                    '' get the constructor argument of the attribute as applied to this member
                    '' to retrieve the value we need to make a decision.
                    Dim i As DesignerSerializationVisibility = cattr.ConstructorArguments(0).Value

                    '' make a decision.
                    If i = DesignerSerializationVisibility.Hidden Then Return VisibilityType.Invisible

                    Return VisibilityType.Visible

                End If
            Next

        End Function

        ''' <summary>
        ''' Gets a value indicating runtime browsability. 1 is visible, -1 is not visible, and 0 is not present.
        ''' </summary>
        ''' <param name="inf"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Browsable(inf As Object) As VisibilityType
            Browsable = VisibilityType.Absent

            Dim ift As IEnumerable(Of CustomAttributeData) = Nothing
            Try
                Dim mbt As MemberInfo = TryCast(inf, MemberInfo)
                Select Case mbt.MemberType

                    Case MemberTypes.Field
                        ''unbox
                        ift = CType(inf, FieldInfo).CustomAttributes

                    Case MemberTypes.Property
                        ''unbox
                        ift = CType(inf, PropertyInfo).CustomAttributes

                    Case Else

                End Select

            Catch ex As Exception
                '' get the other kind of object's attributes
                ift = inf.GetType.CustomAttributes
            End Try

            If ift Is Nothing Then Return False

            For Each cattr In ift
                If cattr.AttributeType = GetType(BrowsableAttribute) Then
                    '' get the constructor argument of the attribute as applied to this member
                    '' to retrieve the value we need to make a decision.
                    Dim i As Boolean = cattr.ConstructorArguments(0).Value

                    '' make a decision.
                    If i Then Return VisibilityType.Visible

                    Return VisibilityType.Invisible
                End If
            Next

        End Function

        Public Function GetPropertyDescription(inf As Object) As String

            If inf.GetType() = GetType(PropertyDescriptor) Then

                Dim s As String = CType(inf, PropertyDescriptor).Description
                If String.IsNullOrWhiteSpace(s) Then
                    Return CType(inf, PropertyDescriptor).DisplayName
                End If

            Else

                Dim pe As PropertyInfo = CType(inf, PropertyInfo)

                For Each cattr In pe.CustomAttributes

                    If cattr.AttributeType = GetType(DescriptionAttribute) Then
                        Return cattr.ConstructorArguments(0).Value
                    End If
                Next

                Return pe.Name
            End If

            Return Nothing
        End Function

        Public Function WriteJavaScriptObjectGraphTemplate(t As System.Type, Optional instance As Object = Nothing, Optional flags As SerializerAdapterFlags = SerializerAdapterFlags.NoFields Or SerializerAdapterFlags.DataMember Or SerializerAdapterFlags.MatchPublic) As String

            Dim mp() As MemberInfo = Nothing
            Dim mi() As MemberInfo = ISAServices.MakeSerializeMap(flags, t, mp)
            Dim fe As FieldInfo
            Dim pe As PropertyInfo

            Dim sb As New StringBuilder
            Dim mt As System.Type = Nothing
            Dim n As String = ""

            If mi Is Nothing Then Return ""

            sb.AppendLine("function " & t.Name & "() {")
            sb.AppendLine("")

            For Each mbr In mi

                Select Case mbr.MemberType

                    Case MemberTypes.Field
                        fe = mbr

                        If Not fe.FieldType.IsValueType AndAlso instance IsNot Nothing AndAlso Not IsNothing(fe.GetValue(instance)) Then
                            mt = fe.GetValue(instance).GetType
                        Else
                            mt = fe.FieldType
                        End If

                        n = fe.Name
                    Case MemberTypes.Property

                        pe = mbr

                        If Not pe.PropertyType.IsValueType AndAlso instance IsNot Nothing AndAlso Not IsNothing(pe.GetValue(instance)) Then
                            mt = pe.GetValue(instance).GetType
                        Else
                            mt = pe.PropertyType
                        End If

                        n = pe.Name

                End Select

                If mt = GetType(Guid) Then
                    sb.AppendLine("    " & "this." & n & " = """ & Guid.Empty.ToString & """;")
                ElseIf mt = GetType(Boolean) Then
                    sb.AppendLine("    " & "this." & n & " = false;")
                ElseIf mt.GetInterface("ICollection") IsNot Nothing Then
                    Dim st As String = mt.Name.Replace("[]", "")

                    If mt.GetInterface("IList") IsNot Nothing Then

                        Dim x As PropertyInfo = mt.GetProperty("Item")
                        If x IsNot Nothing Then st = x.PropertyType.Name

                    End If


                    sb.AppendLine("    " & "this." & n & " = [];  // array of " & st)
                ElseIf mt = GetType(String) Then
                    sb.AppendLine("    " & "this." & n & " = """";")
                ElseIf mt.IsPrimitive Then
                    sb.AppendLine("    " & "this." & n & " = 0;")
                Else
                    sb.AppendLine("    " & "this." & n & " = null; // instance of " & mt.Name)
                End If
            Next

            sb.AppendLine("")
            sb.AppendLine("}")
            sb.AppendLine("")

            Dim o As Object

            For Each mbr In mi
                o = Nothing

                Select Case mbr.MemberType

                    Case MemberTypes.Field
                        fe = mbr

                        If Not fe.FieldType.IsValueType AndAlso instance IsNot Nothing AndAlso Not IsNothing(fe.GetValue(instance)) Then
                            o = fe.GetValue(instance)
                            mt = o.GetType
                        Else
                            mt = fe.FieldType
                        End If

                        n = fe.Name

                    Case MemberTypes.Property
                        pe = mbr

                        If Not pe.PropertyType.IsValueType AndAlso instance IsNot Nothing AndAlso Not IsNothing(pe.GetValue(instance)) Then
                            o = pe.GetValue(instance)
                            mt = o.GetType
                        Else
                            mt = pe.PropertyType
                        End If

                        n = pe.Name

                End Select

                If o IsNot Nothing Then
                    If o.GetType.GetInterface("ICollection") IsNot Nothing Then
                        If CType(o, ICollection).Count = 0 Then Continue For
                        o = CType(o, ICollection)(0)
                    End If

                    sb.Append(WriteJavaScriptObjectGraphTemplate(o.GetType, o, flags))
                End If
            Next

            Return sb.ToString
        End Function

        ''' <summary>
        ''' Calls the internal FormatterServices or Activator classes to create an uninitialized instance of an object.
        ''' </summary>
        ''' <param name="TypeString">The name of the type of object to create.  The name must be fully qualfied or it must be resolvable within the context of this assembly.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetInstance(TypeString As String, Optional t() As System.Type = Nothing) As Object

            If t IsNot Nothing Then
                For Each ti In t
                    If ti.AssemblyQualifiedName = TypeString Then
                        Return GetInstance(ti)
                    End If
                Next
            End If

            Dim Type As System.Type = System.Type.GetType(TypeString)
            Return GetInstance(Type)
        End Function

        ''' <summary>
        ''' Calls the internal FormatterServices or Activator classes to create an uninitialized instance of an object.
        ''' </summary>
        ''' <param name="Type">The type of object to create.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetInstance(Type As System.Type, Optional t() As System.Type = Nothing) As Object

            Dim ti As System.Type = Nothing

            If t IsNot Nothing Then
                For Each ti In t
                    If ti.AssemblyQualifiedName = Type.AssemblyQualifiedName Then
                        Exit For
                    End If
                    ti = Nothing
                Next
            End If

            If ti Is Nothing Then ti = Type

            If Type = GetType(String) Then Return ""

            If ti.GetConstructor(BindingFlags.Instance Or BindingFlags.Public, Nothing, {}, {}) Is Nothing Then
                Try
                    GetInstance = FormatterServices.GetUninitializedObject(ti)
                Catch ex As exception
                    GetInstance = Activator.CreateInstance(ti)
                End Try

            Else
                GetInstance = Activator.CreateInstance(ti)
            End If

            'Try
            '    GetInstance = Activator.CreateInstance(ti)
            'Catch ex As Exception
            '    GetInstance = FormatterServices.GetSafeUninitializedObject(ti)
            'End Try

        End Function

        Public Function PropsWithValues(obj As Object) As List(Of KeyValuePair(Of String, Object))

            Dim pi() As PropertyInfo = obj.GetType.GetProperties
            Dim l As New List(Of KeyValuePair(Of String, Object))
            Dim kv As KeyValuePair(Of String, Object)

            For Each pe In pi
                kv = New KeyValuePair(Of String, Object)(pe.Name, pe.GetValue(obj))
                l.Add(kv)
            Next

            Return l
        End Function

    End Module

#End Region

#Region "Xml serializer"

    Public NotInheritable Class XmlSerializerAdapter
        Implements ISerializerAdapter

        Public Event InstanceHelper(sender As Object, e As InstanceHelperEventArgs) Implements ISerializerAdapter.InstanceHelper

        <ThreadStatic>
        Private _XDoc As XmlDocument

        <ThreadStatic>
        Private _XType As XmlDocumentType

        <ThreadStatic>
        Private _XCurrent As XmlNode

        ''' <summary>
        ''' Sets the format for DateTime object serialization in the XML document.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property DateTimePattern As DateTimePatterns = DateTimePatterns.WindowsVerbose

        ''' <summary>
        ''' Gets or sets the option flags used to determine the seriailzation pattern.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Options As SerializerAdapterFlags = SerializerAdapterFlags.MatchPrivate Or SerializerAdapterFlags.MatchPublic Or SerializerAdapterFlags.MatchBrowsable Or SerializerAdapterFlags.MatchReadProperty Or SerializerAdapterFlags.MatchWriteProperty Implements ISerializerAdapter.Options

        Public Sub New()
            Reset()
        End Sub

        ''' <summary>
        ''' Reset the serializer to a just-create state, clearing all buffers and work items.
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Reset() Implements ISerializerAdapter.Reset
            _XDoc = New XmlDocument
            _XType = _XDoc.CreateDocumentType("AdaptableSerializer2Document", Nothing, Nothing, Nothing)

            _XDoc.AppendChild(_XType)
        End Sub

        ''' <summary>
        ''' Commit an object tree to a file through serialization.
        ''' </summary>
        ''' <param name="obj">The root of the object tree to commit.</param>
        ''' <param name="fileName">The full name of the to write.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Commit(obj As Object, fileName As String) As Boolean Implements ISerializerAdapter.Commit
            Reset()
            SerializeObject(obj)
            Return Commit(fileName)
        End Function

        ''' <summary>
        ''' Commit an object tree to a stream through serialization.
        ''' </summary>
        ''' <param name="obj">The root of the object tree to commit.</param>
        ''' <param name="stream">The stream to which to commit the object.</param>
        ''' <param name="closeStream">Inidicates whether to leave the stream open after commit.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Commit(obj As Object, stream As IO.Stream, Optional closeStream As Boolean = True) As Boolean Implements ISerializerAdapter.Commit
            Reset()
            SerializeObject(obj)
            Return Commit(stream)
        End Function

        ''' <summary>
        ''' Internal function to commit the present work document to the file.
        ''' </summary>
        ''' <param name="fileName">The name of the file to which to write.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Commit(fileName As String) As Boolean
            Try
                Dim fs As New FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None)
                Return Commit(fs, True)
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Internal function to commit the present work document to the file.
        ''' </summary>
        ''' <param name="stream">The stream to which to write.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Commit(stream As System.IO.Stream, Optional ByVal closeStream As Boolean = True) As Boolean
            Try
                Dim xws As New XmlWriterSettings

                xws.CheckCharacters = True
                xws.Indent = True
                xws.IndentChars = "    "
                xws.NamespaceHandling = NamespaceHandling.OmitDuplicates
                xws.NewLineChars = vbCrLf
                xws.NewLineHandling = NewLineHandling.None
                xws.NewLineOnAttributes = False
                xws.CloseOutput = False
                xws.Encoding = UnicodeEncoding.Unicode

                Dim xw As XmlWriter = XmlWriter.Create(stream, xws)

                _XDoc.WriteContentTo(xw)
                xw.Close()

                If closeStream Then stream.Close() Else stream.Seek(0, SeekOrigin.Begin)

            Catch ex As Exception
                Return False
            End Try

            Return True
        End Function

        ''' <summary>
        ''' Decommit an object tree from a file through deserialization.
        ''' </summary>
        ''' <param name="fileName">The full name of the file to read.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Decommit(fileName As String) As Object Implements ISerializerAdapter.Decommit
            Dim fs As New FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Return Decommit(fs)
        End Function

        ''' <summary>
        ''' Decommit an object tree from a stream.
        ''' </summary>
        ''' <param name="stream">The stream from which to read.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Decommit(stream As System.IO.Stream) As Object Implements ISerializerAdapter.Decommit
            Dim setting As Object = Nothing
            If Decommit(stream, setting) Then
                Return setting
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' Decommit serialized data to a root object that may or not be preinstantiated.
        ''' </summary>
        ''' <param name="fileName">The full name of the file to read.</param>
        ''' <param name="rootObject">The object to receive the decommited data.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Decommit(fileName As String, ByRef rootObject As Object) As Boolean Implements ISerializerAdapter.Decommit
            Dim fs As New FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Return Decommit(fs, rootObject)
        End Function

        ''' <summary>
        ''' Decommit serialized data to a root object that may or not be preinstantiated.
        ''' </summary>
        ''' <param name="stream">The stream from which to read.</param>
        ''' <param name="rootObject">The object to receive the decommited data.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Decommit(stream As IO.Stream, ByRef rootObject As Object) As Boolean Implements ISerializerAdapter.Decommit
            Dim b() As Byte

            ReDim b(stream.Length - 1)
            stream.Read(b, 0, stream.Length)
            stream.Close()

            _XDoc = New XmlDocument
            _XDoc.InnerXml = DataTools.ByteOrderMark.SafeTextRead(b)

            DeserializeObject(rootObject, _XDoc.ChildNodes(2))
            Decommit = (rootObject IsNot Nothing)
        End Function

        ''' <summary>
        ''' Internally serialize an object to the working Xml document.
        ''' </summary>
        ''' <param name="Setting"></param>
        ''' <param name="NodeName"></param>
        ''' <remarks></remarks>
        Private Sub SerializeObject(Setting As Object, Optional NodeName As String = Nothing)

            '' Create the Xml element.
            Dim xn As XmlNode = _XDoc.CreateElement(If(NodeName Is Nothing, Setting.GetType.Name, NodeName))

            Dim mp() As PropertyInfo = Nothing

            '' Make the serializer map.
            Dim mi() As MemberInfo

            mi = Cache.FindMembers(Setting.GetType)

            If mi Is Nothing Then
                mi = MakeSerializeMap(_Options, Setting.GetType, mp)
                Cache.Add(Setting.GetType, _Options, mi)
            End If

            Dim xa As XmlAttribute

            '' see if this type is composite and contains a Name attribute.
            If Setting.GetType.GetProperty("Name") IsNot Nothing Then
                '' apply the name attribute to the node.
                xa = _XDoc.CreateAttribute("Name")
                xa.Value = Setting.Name
                xn.Attributes.Append(xa)
            End If

            '' get the type converter for the object type.
            '' if we can just serialize to string, then we don't have to go any further.
            Dim objConv As TypeConverter = TypeDescriptor.GetConverter(Setting.GetType)

            If objConv.CanConvertFrom(GetType(String)) AndAlso
                objConv.CanConvertTo(GetType(String)) Then

                If Setting.GetType() = GetType(Date) Then
                    '' encode date time according to the selected pattern.
                    xn.InnerXml = CDate(Setting).ToString(CStr(ChrW(DateTimePattern)))
                Else
                    xn.InnerXml = objConv.ConvertTo(Setting, GetType(String))
                End If

                '' append the child to the parent node and return.
                If _XCurrent Is Nothing Then
                    _XDoc.AppendChild(xn)
                Else
                    _XCurrent.AppendChild(xn)
                End If

                Return
            ElseIf mi Is Nothing OrElse mi.Count = 0 Then
                '' we cannot convert the value to a string and this is not a composite object
                '' therefore the property is unserializable even though it passed the
                '' constraints.
                If Setting.GetType.GetInterface("IList`1") Is Nothing AndAlso Setting.GetType.GetInterface("IList") Is Nothing AndAlso Setting.GetType.IsArray = False Then Return
            End If

            '' see if the current node is nothing, append to the document,
            '' otherwise, append to the current node.
            If _XCurrent Is Nothing Then
                _XDoc.AppendChild(xn)
            Else
                _XCurrent.AppendChild(xn)
            End If

            '' now the current node is this node as we go composite.
            _XCurrent = xn

            '' create the verbose type info for the current node because it is composite
            '' and will need to be instantiated.
            xa = _XDoc.CreateAttribute("Type")
            xa.Value = Setting.GetType.AssemblyQualifiedName
            xn.Attributes.Append(xa)

            '' if this node is a collection we need to record the collection items:
            If Setting.GetType.GetInterface("IList") IsNot Nothing OrElse Setting.GetType.GetInterface("IList`1") IsNot Nothing OrElse Setting.GetType.IsArray = True Then

                Dim ic As Object,
                    x As Integer = 0


                For Each ic In CType(Setting, IEnumerable)
                    SerializeObject(ic, "Element")
                    x += 1
                Next

                xa = _XDoc.CreateAttribute("Count")
                xa.Value = x
                xn.Attributes.Append(xa)

            End If

            '' now iterate through the properties and fields and get their values.
            Dim fe As FieldInfo = Nothing
            Dim pe As PropertyInfo = Nothing
            Dim pr As PropertyInfo = Nothing

            If mi Is Nothing Then
                _XCurrent = _XCurrent.ParentNode
                Return
            End If

            For Each mbr In mi

                '' properties
                If mbr.MemberType = MemberTypes.Property Then
                    pe = mbr

                    '' we cannot serialize index properties.
                    '' however if they are backed by an array that
                    '' array can be serialized.
                    If pe.GetIndexParameters.Count > 0 Then Continue For

                    Try
                        Dim o As Object = pe.GetValue(Setting, Nothing)
                        '' recursively call SerializeObject to descend the object tree.
                        If o IsNot Nothing Then SerializeObject(o, pe.Name)
                    Catch ex As Exception

                    End Try
                ElseIf mbr.MemberType = MemberTypes.Field Then
                    '' fields
                    fe = mbr

                    Try
                        Dim o As Object = fe.GetValue(Setting)
                        '' recursively call SerializeObject to descend the object tree.

                        If o IsNot Nothing Then

                            '' check if we need to replace the property name
                            If Options And SerializerAdapterFlags.MatchUsePropName Then
                                pr = MatchFieldToProperty(fe.Name, mp)
                                If pr IsNot Nothing Then
                                    SerializeObject(o, pr.Name)
                                End If
                            End If

                            If pr Is Nothing Then SerializeObject(o, fe.Name) Else pr = Nothing

                        End If

                    Catch ex As Exception

                    End Try
                End If
            Next

            '' finished with this call, retrace back up the tree.
            _XCurrent = _XCurrent.ParentNode

        End Sub

        ''' <summary>
        ''' Internally deserialize an object from the Xml document.
        ''' </summary>
        ''' <param name="Setting"></param>
        ''' <param name="xml"></param>
        ''' <param name="fieldType"></param>
        ''' <remarks></remarks>
        Private Sub DeserializeObject(ByRef Setting As Object, xml As XmlNode, Optional fieldType As System.Type = Nothing)

            Dim tn As String

            '' perform tasks to get a property Type class and instantiate a blank object.
            If Setting Is Nothing AndAlso xml.Attributes("Type") IsNot Nothing Then
                tn = xml.Attributes("Type").Value

                If fieldType.IsArray Then

                    Setting = Array.CreateInstance(fieldType.GetElementType, xml.ChildNodes.Count)
                Else
                    Try
                        If Setting Is Nothing Then Setting = GetInstance(tn, _KnownTypes)
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try

                    If Setting Is Nothing Then Return
                End If
            End If

            If Setting IsNot Nothing AndAlso fieldType Is Nothing Then fieldType = Setting.GetType

            '' if after all that we still do not have a fieldType we have failed to
            '' retrieve the type information and/or create an object.
            '' we cannot continue.
            If fieldType Is Nothing Then Return

            '' get the serialization map according to the Options specified.
            '' the options must be identical to the options used to serialize
            '' the file or it will not deserialize correctly.
            Dim mp() As PropertyInfo = Nothing
            Dim mi() As MemberInfo


            mi = Cache.FindMembers(fieldType)

            If mi Is Nothing Then
                mi = MakeSerializeMap(_Options, fieldType, mp)
                Cache.Add(fieldType, _Options, mi)
            End If

            Dim objConv As TypeConverter = TypeDescriptor.GetConverter(fieldType)

            If objConv.CanConvertFrom(GetType(String)) AndAlso _
                objConv.CanConvertTo(GetType(String)) Then

                Setting = objConv.ConvertFrom(xml.InnerXml)

                Return
            ElseIf mi Is Nothing OrElse mi.Count = 0 Then
                If Setting.GetType.GetInterface("IList") Is Nothing AndAlso Setting.GetType.IsArray = False Then Return
            End If

            If Setting.GetType.GetInterface("IList") IsNot Nothing OrElse Setting.GetType.IsArray = True Then

                Dim x As Integer = 0,
                    c As Integer = CInt(xml.Attributes("Count").Value) - 1

                Dim subSetting As Object = Nothing

                For x = 0 To c
                    subSetting = Nothing
                    DeserializeObject(subSetting, xml.ChildNodes(x), Setting.GetType.GetElementType)

                    If Setting.GetType.IsArray AndAlso CType(Setting, Object()).Count >= x Then
                        Setting(x) = subSetting
                    Else
                        If subSetting IsNot Nothing Then CType(Setting, IList).Add(subSetting)
                    End If
                Next

            End If

            Dim fe As FieldInfo = Nothing
            Dim pe As PropertyInfo = Nothing
            Dim pr As PropertyInfo = Nothing

            If mi Is Nothing Then Return

            For Each mbr In mi

                If mbr.MemberType = MemberTypes.Property Then
                    pe = mbr
                    Try
                        Dim o As Object = Nothing
                        Dim xn = findNode(pe.Name, xml)

                        If xn IsNot Nothing Then
                            o = pe.GetValue(Setting, Nothing)
                            DeserializeObject(o, xn, pe.PropertyType)
                            pe.SetValue(Setting, o, Nothing)
                        End If
                    Catch ex As Exception

                    End Try

                ElseIf mbr.MemberType = MemberTypes.Field Then
                    fe = mbr

                    Try
                        '' check for associated deserialization flags.
                        If Options And SerializerAdapterFlags.MatchUsePropName Then
                            pr = MatchFieldToProperty(fe.Name, mp)
                        End If

                        Dim o As Object = Nothing
                        Dim xn As XmlNode = Nothing

                        If pr IsNot Nothing Then
                            xn = findNode(pr.Name, xml)
                        Else
                            xn = findNode(fe.Name, xml)
                        End If

                        If xn Is Nothing Then Continue For

                        o = fe.GetValue(Setting)
                        DeserializeObject(o, xn, fe.FieldType)

                        fe.SetValue(Setting, o)
                    Catch ex As Exception

                    End Try

                End If

            Next

        End Sub

        ''' <summary>
        ''' Gets the serialization type implemented by the class.  Text = 0, Binary = 1, Xml = 2
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Type As SerializationTypes Implements ISerializerAdapter.Type
            Get
                Return SerializationTypes.Xml
            End Get
        End Property

        Private Function findNode(name As String, xin As XmlNode) As XmlNode
            For Each xn As XmlNode In xin.ChildNodes
                If xn.Name = name Then Return xn
            Next
            Return Nothing
        End Function

        Public Property KnownTypes As Type() Implements ISerializerAdapter.KnownTypes

        Public Property FileCode As Long Implements ISerializerAdapter.FileCode

    End Class

#End Region

#Region "Binary Structures"

    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Structure BIN_SER_HDR
        Public HeaderSize As Integer
        Public FileCode As Long
        Public VersionMajor As Short
        Public VersionMinor As Short
        Public Timestamp As Date

        Public Sub New(fcVal As String)
            FileCode = FCFromString(fcVal)
            Timestamp = Now
        End Sub

        Public Function FormatStruct() As Blob

            FormatStruct = New Blob With {.InBufferMode = True, .AutoDouble = False, .BufferExtend = 1024 * 1024}

            HeaderSize = Marshal.SizeOf(Me)
            FormatStruct &= StructToBytes(Me)
        End Function

        Public Shared Function DeformatStruct(ByRef bl As Blob, Optional ByRef pos As Integer = -1) As BIN_SER_HDR
            Dim bfmt As New BIN_SER_FMT

            'pos = 0

            If pos <> -1 Then bl.ClipSeek(pos)
            pos = bl.ClipNext

            BytesToStruct(bl.Clip(Marshal.SizeOf(DeformatStruct), BlobTypes.Byte), DeformatStruct)
            pos += DeformatStruct.HeaderSize
            bl.ClipSeek(pos)
        End Function

        Public Shared Function FCFromString(value As String) As Long
            Dim b() As Byte = UTF8Encoding.UTF8.GetBytes(value)
            Dim bl As Blob = b
            FCFromString = bl.Clip(8, BlobTypes.Long)
            bl.Dispose()
        End Function

    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Structure BIN_SER_FMT
        Public DataType As BlobTypes
        Public TypeString As String
        Public Data() As Byte
        Public Name As String
        Public HasChildren As Byte
        Public ChildCount As Integer
        Public Children() As BIN_SER_FMT

        Public Shared Converter As New BlobConverter

        Public Function ChildByName(name As String) As BIN_SER_FMT
            If Children Is Nothing Then Return Nothing

            For Each ch In Children
                If ch.Name = name Then Return ch
            Next

            Return Nothing
        End Function

        Public Function HasChild(name As String) As Boolean
            If Children Is Nothing Then Return False
            For Each ch In Children
                If ch.Name = name Then Return True
            Next

            Return False
        End Function

        Public Function FormatStruct(Optional bl As Blob = Nothing) As Blob
            'Dim bl As Blob

            If bl Is Nothing Then
                bl = New Blob With {.InBufferMode = True, .AutoDouble = False, .BufferExtend = (1024 * 1024)}
            End If

            bl &= If(Data IsNot Nothing, CInt(DataType), BlobTypes.Invalid)

            If Name Is Nothing AndAlso Data Is Nothing AndAlso HasChildren = 0 Then Return bl
            bl.BlobType = BlobTypes.Byte

            bl.StringCatNoNull = True

            If DataType = (BlobTypes.Blob) Then

                If Data IsNot Nothing AndAlso Data.Length > 0 Then
                    bl &= CInt(Data.Length)
                    bl &= Data
                Else
                    bl &= CInt(0)
                    ChildCount = 0
                End If

            ElseIf (DataType And (BlobOrdinalTypes.Invalid)) OrElse Data Is Nothing Then

                If DataType = BlobTypes.Boolean Then
                    bl &= Data(0)
                Else
                    If TypeString IsNot Nothing Then
                        bl &= CShort(TypeString.Length)
                        bl &= TypeString
                    Else
                        bl &= CShort(0)
                    End If

                    If Data IsNot Nothing Then
                        bl &= CInt(Data.Length)
                        bl &= Data
                    Else
                        bl &= CInt(0)
                    End If
                End If
            Else
                bl &= Data
            End If

            If Name Is Nothing Then
                bl &= CShort(0)
            Else
                bl &= CShort(Name.Length)
                bl &= Name
            End If

            bl &= CByte(HasChildren And 1)
            Dim x As Integer = 0

            If HasChildren <> 0 Then
                Dim bfmt As BIN_SER_FMT
                bl &= ChildCount

                For x = 0 To ChildCount - 1
                    bfmt = Children(x)
                    bfmt.FormatStruct(bl)
                Next

            End If

            If bl.BlobType <> BlobTypes.Byte Then bl.BlobType = BlobTypes.Byte
            FormatStruct = bl
        End Function

        Public Shared Function DeformatStruct(ByRef bl As Blob, Optional ByRef pos As Integer = -1) As BIN_SER_FMT
            Dim bfmt As New BIN_SER_FMT
            Dim ln As Integer

            'pos = 0

            If pos <> -1 Then bl.ClipSeek(pos) Else pos = 0
            pos = bl.ClipNext

            bfmt.DataType = CType(bl.Clip(4, BlobTypes.Integer), BlobTypes)
            pos += 4

            If bfmt.DataType = (BlobTypes.Blob) Then

                ln = bl.Clip(4, BlobTypes.Integer)
                pos += 4

                If ln Then

                    bfmt.Data = bl.Clip(ln, BlobTypes.Byte)
                    pos += bfmt.Data.Length
                End If

            ElseIf (bfmt.DataType = BlobTypes.Boolean) Then
                ReDim bfmt.Data(0)
                bfmt.Data(0) = bl.Clip(1, BlobTypes.Byte)
                pos += 1
                '            ElseIf bfmt.DataType = BlobTypes.Image Then

            ElseIf (bfmt.DataType And (BlobOrdinalTypes.Invalid)) Then
                ln = bl.Clip(2, BlobTypes.Short)

                If ln Then
                    bfmt.TypeString = bl.Clip(2 * ln, BlobTypes.String)
                End If

                pos += 2 + (ln * 2)

                ln = bl.Clip(4, BlobTypes.Integer)
                pos += 4

                If ln Then

                    bfmt.Data = bl.Clip(ln, BlobTypes.Byte)
                    pos += bfmt.Data.Length

                End If
            Else
                If bfmt.DataType = BlobTypes.Byte OrElse bfmt.DataType = BlobTypes.SByte Then
                    bfmt.Data = {bl.Clip(1, BlobTypes.Byte)}
                    pos += 1
                Else
                    bfmt.Data = bl.Clip(Blob.BlobTypeSize(bfmt.DataType), BlobTypes.Byte)
                    pos += Blob.BlobTypeSize(bfmt.DataType)
                End If
            End If

            ln = bl.Clip(2, BlobTypes.Short)

            If ln Then
                bfmt.Name = bl.Clip(2 * ln, BlobTypes.String)
            End If

            pos += 2 + (ln * 2)

            bfmt.HasChildren = bl.Clip(1, BlobTypes.Byte)
            pos += 1

            If bfmt.HasChildren <> 0 Then
                bfmt.ChildCount = CInt(bl.Clip(4, BlobTypes.Integer))
                pos += 4

                If bfmt.ChildCount >= &H7FFFFFFE Then
                    MsgBox("Yo!")
                End If

                ReDim bfmt.Children(bfmt.ChildCount - 1)

                For i = 0 To bfmt.ChildCount - 1
                    bfmt.Children(i) = BIN_SER_FMT.DeformatStruct(bl, pos)
                Next

            End If

            DeformatStruct = bfmt
        End Function

        <ThreadStatic>
        Private Shared blstr As MemPtr

        <ThreadStatic>
        Private Shared bllen As Long = 0

        Public Property Value() As Object
            Get
                If Data Is Nothing Then Return Nothing

                Dim bl As Blob = Data
                Value = bl.Clip(0, Data.Length, DataType)
                bl.Dispose()
            End Get
            Set(value As Object)
                Dim bl As Blob
                If Blob.TypeToBlobType(value.GetType) = BlobTypes.Invalid Then
                    Throw New ArgumentException(value.GetType.FullName & " is not supported by serialization")
                    Return
                Else
                    Dim mm As MemPtr
                    Dim i As Integer

                    If value Is Nothing Then
                        Data = Nothing
                        Return
                    End If

                    Select Case value.GetType

                        Case GetType(Byte())
                            Me.DataType = (BlobTypes.Blob)
                            i = CType(value, Byte()).Length

                            ReDim Data(i - 1)
                            MemCpy(Data, CType(value, Byte()), i)
                            Return
                        Case GetType(String)
                            Dim s As String = CStr(value).Trim(ChrW(0))

                            i = (s.Length * 2) + 2
                            Me.DataType = BlobTypes.String

                            If i <= 2 Then
                                Data = Nothing
                                Return
                            End If

                            If i > bllen Then
                                bllen = i
                                blstr.ReAlloc(i)
                            End If

                            blstr.SetString(0, s)

                            Data = blstr.GetByteArray(0, i - 2)
                            blstr.Free()
                            Return

                        Case GetType(Integer)
                            Me.DataType = BlobTypes.Integer
                            Data = BitConverter.GetBytes(CInt(value))
                            Return

                        'Case GetType(Integer())
                        '    Me.DataType = BlobTypes.Integer
                        '    i = CType(value, Short()).Length * 4

                        '    ReDim Data(i - 1)
                        '    MemCpy(Data, CType(value, Integer()), i)
                        '    Return

                        'Case GetType(SByte())
                        '    Me.DataType = BlobTypes.SByte
                        '    i = CType(value, Byte()).Length

                        '    ReDim Data(i - 1)
                        '    MemCpy(Data, CType(value, SByte()), i)
                        '    Return

                        'Case GetType(Char())
                        '    Me.DataType = BlobTypes.Char
                        '    i = CType(value, Short()).Length * 2

                        '    ReDim Data(i - 1)
                        '    MemCpy(Data, CType(value, Char()), i)
                        '    Return

                        'Case GetType(Short())
                        '    Me.DataType = BlobTypes.Short
                        '    i = CType(value, Short()).Length * 2

                        '    ReDim Data(i - 1)
                        '    MemCpy(Data, CType(value, Short()), i)
                        '    Return

                        'Case GetType(UShort())
                        '    Me.DataType = BlobTypes.UShort
                        '    i = CType(value, Short()).Length * 2

                        '    ReDim Data(i - 1)
                        '    MemCpy(Data, CType(value, UShort()), i)
                        '    Return

                        'Case GetType(UInteger())
                        '    Me.DataType = BlobTypes.UInteger
                        '    i = CType(value, Short()).Length * 4

                        '    ReDim Data(i - 1)
                        '    MemCpy(Data, CType(value, UInteger()), i)
                        '    Return

                        'Case GetType(Long())
                        '    Me.DataType = BlobTypes.Long
                        '    i = CType(value, Short()).Length * 8

                        '    ReDim Data(i - 1)
                        '    MemCpy(Data, CType(value, Long()), i)
                        '    Return

                        'Case GetType(ULong())
                        '    Me.DataType = BlobTypes.ULong
                        '    i = CType(value, Short()).Length * 8

                        '    ReDim Data(i - 1)
                        '    MemCpy(Data, CType(value, ULong()), i)
                        '    Return

                        'Case GetType(Single())
                        '    Me.DataType = BlobTypes.Single
                        '    i = CType(value, Short()).Length * 4

                        '    ReDim Data(i - 1)
                        '    MemCpy(Data, CType(value, Single()), i)
                        '    Return

                        'Case GetType(Double())
                        '    Me.DataType = BlobTypes.Double
                        '    i = CType(value, Short()).Length * 8

                        '    ReDim Data(i - 1)
                        '    MemCpy(Data, CType(value, Double()), i)
                        '    Return

                        Case GetType(Boolean)
                            Me.DataType = BlobTypes.Boolean
                            ReDim Data(0)
                            If (CBool(value) = True) Then Data(0) = 1
                            Return

                        Case GetType(SByte)
                            Me.DataType = BlobTypes.SByte
                            Data = BitConverter.GetBytes(CSByte(value))
                            Return

                        Case GetType(Byte)
                            Me.DataType = BlobTypes.Byte
                            Data = BitConverter.GetBytes(CByte(value))
                            Return

                        Case GetType(Short)
                            Me.DataType = BlobTypes.Short
                            Data = BitConverter.GetBytes(CShort(value))
                            Return

                        Case GetType(UShort)
                            Me.DataType = BlobTypes.UShort
                            Data = BitConverter.GetBytes(CUShort(value))
                            Return

                        Case GetType(UInteger)
                            Me.DataType = BlobTypes.UInteger
                            Data = BitConverter.GetBytes(CUInt(value))
                            Return

                        Case GetType(Long)
                            Me.DataType = BlobTypes.Long
                            Data = BitConverter.GetBytes(CLng(value))
                            Return

                        Case GetType(ULong)
                            Me.DataType = BlobTypes.ULong
                            Data = BitConverter.GetBytes(CULng(value))
                            Return


                        Case GetType(Single)
                            Me.DataType = BlobTypes.Single
                            Data = BitConverter.GetBytes(CSng(value))
                            Return

                        Case GetType(Double)
                            Me.DataType = BlobTypes.Double
                            Data = BitConverter.GetBytes(CDbl(value))
                            Return

                        Case GetType(Decimal)
                            Me.DataType = BlobTypes.Decimal
                            mm = CType(value, Decimal)
                            Data = mm
                            mm.Free()
                            Return

                        Case GetType(BigInteger)
                            Me.DataType = BlobTypes.BigInteger
                            mm = CType(value, BigInteger).ToByteArray
                            Data = mm
                            mm.Free()
                            Return

                        Case GetType(Date)
                            Me.DataType = BlobTypes.Date
                            Data = BitConverter.GetBytes(CDate(value).ToBinary)
                            Return

                        Case GetType(Guid)
                            Me.DataType = BlobTypes.Guid
                            Data = CType(value, Guid).ToByteArray
                            Return

                        Case GetType(Color)
                            Me.DataType = BlobTypes.Color
                            mm = CType(value, Color)
                            Data = mm
                            mm.Free()
                            Return

                        Case GetType(Bitmap), GetType(Image)
                            Data = BlobConverter.ImageToBytes(value, Imaging.ImageFormat.Png)
                            Me.DataType = BlobTypes.Image
                            Return

                        Case Else
                            bl = Converter.ConvertFrom(value)
                            Me.DataType = Blob.TypeToBlobType(value.GetType)

                    End Select

                End If

                Data = bl
            End Set
        End Property

        Public Overrides Function ToString() As String
            ToString = If(Value Is Nothing, "", If(TypeOf Value Is String, Value, Value.ToString))
        End Function

    End Structure

#End Region

#Region "Binary Serializer"

    Public Class BinarySerializerAdapter
        Implements ISerializerAdapter

        Public Event InstanceHelper(sender As Object, e As InstanceHelperEventArgs) Implements ISerializerAdapter.InstanceHelper

        Private _Buffer() As Byte
        Private _Header As New BIN_SER_HDR With {.VersionMajor = 2, .VersionMinor = 0, .FileCode = BIN_SER_HDR.FCFromString("ADBNSRX2")}
        Private _Root As BIN_SER_FMT

        Public Property Options As SerializerAdapterFlags = SerializerAdapterFlags.MatchPrivate Or SerializerAdapterFlags.MatchPublic Or SerializerAdapterFlags.MatchBrowsable Or SerializerAdapterFlags.MatchReadProperty Or SerializerAdapterFlags.MatchWriteProperty Implements ISerializerAdapter.Options

        Public Sub New()
            Reset()
        End Sub

        ''' <summary>
        ''' Reset the serializer to a just-create state, clearing all buffers and work items.
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Reset() Implements ISerializerAdapter.Reset
            Dim fc As Long = _Header.FileCode

            _Buffer = Nothing
            _Root = New BIN_SER_FMT
            _Header = New BIN_SER_HDR With {.VersionMajor = 2, .VersionMinor = 0, .FileCode = fc, .Timestamp = Now}
        End Sub

        ''' <summary>
        ''' A 64 bit (8 byte) unique identifier to be included in the binary header.
        ''' This value must be the same on both serialization and deserialization or the function will fail.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property FileCode As Long Implements ISerializerAdapter.FileCode
            Get
                Return _Header.FileCode
            End Get
            Set(value As Long)
                _Header.FileCode = value
            End Set
        End Property

        ''' <summary>
        ''' Private function commit the working structures to a file.
        ''' </summary>
        ''' <param name="fileName">The file to which to write.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Commit(fileName As String) As Boolean
            Try
                Dim fs As New FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None)
                Return Commit(fs, True)
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Private function commit the working structures to a stream.
        ''' </summary>
        ''' <param name="stream">The stream to which to write.</param>
        ''' <param name="closeStream">Speficies whether or not to close the stream when the operation is complete.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Commit(stream As System.IO.Stream, Optional ByVal closeStream As Boolean = True) As Boolean
            Try
                ''Dim by As New Blob With {.AutoDouble = False, .InBufferMode = True, .BufferExtend = 1024 * 1024}
                Dim by As Blob = _Header.FormatStruct


                _Root.FormatStruct(by)
                stream.Write(by.GrabBytes, 0, by.Length)

                by.Dispose()

                If closeStream Then stream.Close() Else stream.Seek(0, SeekOrigin.Begin)
            Catch ex As Exception
                Return False
            End Try

            Return True
        End Function

        ''' <summary>
        ''' Commit an object tree to a file through serialization.
        ''' </summary>
        ''' <param name="obj">The root of the object tree to commit.</param>
        ''' <param name="fileName">The full name of the to write.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Commit(obj As Object, fileName As String) As Boolean Implements ISerializerAdapter.Commit

            Reset()
            Monitor.Enter(obj)
            SerializeObject(obj, _Root)
            Monitor.Exit(obj)

            Try
                Dim by As New Blob With {.InBufferMode = True, .BufferExtend = 2 * 1024 * 1024}

                by &= _Header.FormatStruct().GrabBytes
                _Root.FormatStruct(by)

                by.Write(fileName)
                by.Dispose()


                'Dim ry = _Header.FormatStruct
                'Dim by = _Root.FormatStruct

                'Dim b2new As New Blob

                'b2new.Length = by.Length + ry.Length

                'MemCpy(b2new.DangerousGetHandle, ry.DangerousGetHandle, ry.Length)
                'MemCpy(b2new.DangerousGetHandle + ry.Length, by.DangerousGetHandle, by.Length)

                'b2new.Write(fileName)

                'by.Dispose()
                'ry.Dispose()
                'b2new.Dispose()



            Catch ex As Exception
                Return False
            End Try

            Return True
        End Function

        ''' <summary>
        ''' Commit an object tree to a stream through serialization.
        ''' </summary>
        ''' <param name="obj">The root of the object tree to commit.</param>
        ''' <param name="stream">The stream to which to commit the object.</param>
        ''' <param name="closeStream">Specifies whether to close the stream when the operation is complete.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Commit(obj As Object, stream As IO.Stream, Optional closeStream As Boolean = True) As Boolean Implements ISerializerAdapter.Commit
            Reset()
            Monitor.Enter(obj)
            SerializeObject(obj, _Root)
            Monitor.Exit(obj)
            Return Commit(stream, closeStream)
        End Function

        ''' <summary>
        ''' Decommit an object tree from a file through deserialization.
        ''' </summary>
        ''' <param name="fileName">The full name of the file to read.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Decommit(fileName As String) As Object Implements ISerializerAdapter.Decommit
            Decommit = Nothing
            Decommit(fileName, Decommit)
            'Dim fs As New FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            'Return Decommit(fs)
        End Function

        ''' <summary>
        ''' Decommit an object tree from a stream.
        ''' </summary>
        ''' <param name="stream">The stream from which to read.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Decommit(stream As System.IO.Stream) As Object Implements ISerializerAdapter.Decommit
            Dim setting As Object = Nothing
            If Decommit(stream, setting) Then
                Return setting
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' Decommit serialized data to a root object that may or not be preinstantiated.
        ''' </summary>
        ''' <param name="fileName">The full name of the file to read.</param>
        ''' <param name="rootObject">The object to receive the decommited data.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Decommit(fileName As String, ByRef rootObject As Object) As Boolean Implements ISerializerAdapter.Decommit
            Dim bl As New Blob

            'bl.CreatePrivateHeap()
            bl.Read(fileName)

            _Header = BIN_SER_HDR.DeformatStruct(bl)
            _Root = BIN_SER_FMT.DeformatStruct(bl)

            DeserializeObject(rootObject, _Root)
            Decommit = (rootObject IsNot Nothing)

            bl.Dispose()

            'Dim fs As New FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            'Return Decommit(fs, rootObject)
        End Function

        ''' <summary>
        ''' Decommit serialized data to a root object that may or not be preinstantiated.
        ''' </summary>
        ''' <param name="stream">The stream from which to read.</param>
        ''' <param name="rootObject">The object to receive the decommited data.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Decommit(stream As IO.Stream, ByRef rootObject As Object) As Boolean Implements ISerializerAdapter.Decommit
            Dim b() As Byte
            Dim bl As Blob = Nothing
            ReDim b(stream.Length - 1)
            stream.Read(b, 0, stream.Length)
            stream.Close()
            _Buffer = b
            bl &= _Buffer

            _Header = BIN_SER_HDR.DeformatStruct(bl)
            _Root = BIN_SER_FMT.DeformatStruct(bl)

            DeserializeObject(rootObject, _Root)
            Decommit = (rootObject IsNot Nothing)
        End Function

        ''' <summary>
        ''' Internally serialize an object to a binary structure.
        ''' </summary>
        ''' <param name="Setting"></param>
        ''' <param name="bfmt"></param>
        ''' <remarks></remarks>
        Private Sub SerializeObject(Setting As Object, ByRef bfmt As BIN_SER_FMT)

            If Setting Is Nothing Then Return

            Dim mp() As MemberInfo = Nothing
            Dim mi() As MemberInfo

            Dim sgt As Type = Setting.GetType

            If TypeOf Setting Is String Then
                bfmt.Value = CStr(Setting)
            End If

            If Blob.TypeToBlobType(sgt) = (BlobTypes.Invalid) Then
                bfmt.HasChildren = 0

                mi = Cache.FindMembers(sgt)

                If mi Is Nothing Then
                    mi = MakeSerializeMap(_Options, sgt, mp)
                    If mi IsNot Nothing Then Cache.Add(sgt, _Options, mi)
                End If

                Dim isso As ISelfSerializingObject = If(sgt.GetInterface("ISelfSerializingObject") IsNot Nothing, Setting, Nothing)

                If (isso IsNot Nothing) Then
                    bfmt.DataType = BlobTypes.Invalid
                    bfmt.TypeString = sgt.AssemblyQualifiedName

                    bfmt.Data = isso.Serialize
                    Return
                End If

                'Dim objConv As TypeConverter = TypeDescriptor.GetConverter(sgt)

                'If objConv.CanConvertFrom(GetType(String)) AndAlso
                '        objConv.CanConvertTo(GetType(String)) Then

                '    bfmt.Value = objConv.ConvertTo(Setting, GetType(String))
                '    Return
                'Else

                If (mi Is Nothing OrElse mi.Count = 0) AndAlso (sgt.GetInterface("ICollection") Is Nothing) AndAlso (sgt.GetInterface("ICollection`1") Is Nothing) Then
                    Return
                End If

            Else
                bfmt.HasChildren = 0
                bfmt.Value = Setting
                Return
            End If

            bfmt.TypeString = sgt.AssemblyQualifiedName
            bfmt.HasChildren = 1

            Dim x As Integer = 0

            If sgt.GetInterface("ICollection") IsNot Nothing OrElse sgt.GetInterface("ICollection`1") IsNot Nothing Then

                If Setting.GetType.IsArray Then
                    bfmt.ChildCount = Setting.Length
                Else
                    bfmt.ChildCount = Setting.Count
                End If

                ReDim bfmt.Children(bfmt.ChildCount - 1)

                For Each ic In Setting
                    If ic Is Nothing Then Continue For
                    SerializeObject(ic, bfmt.Children(x))
                    x += 1
                Next

                If x <> bfmt.Children.Length Then
                    ReDim Preserve bfmt.Children(x - 1)
                    bfmt.ChildCount = x
                End If

                Return
            End If

            Dim fe As FieldInfo = Nothing
            Dim pe As PropertyInfo = Nothing
            Dim pr As PropertyInfo = Nothing

            If mi Is Nothing Then Return

            Dim o As Object

            For Each mbr In mi

                o = Nothing

                If mbr.MemberType = MemberTypes.Property Then
                    pe = mbr

                    '' we cannot serialize index properties.
                    '' however if they are backed by an array that
                    '' array can be serialized.
                    If pe.GetIndexParameters.Count > 0 Then Continue For

                    Try
                        o = pe.GetValue(Setting, Nothing)
                        If o Is Nothing AndAlso pe.PropertyType = GetType(String) Then o = ""

                        If o IsNot Nothing Then
                            ReDim Preserve bfmt.Children(x)
                            bfmt.Children(x).Name = pe.Name
                            SerializeObject(o, bfmt.Children(x))

                            x += 1
                        End If

                    Catch ex As Exception
                        MsgBox(String.Format("Error with property {0} of classtype {1}: {2}" & vbCrLf & vbCrLf & ex.StackTrace, pe.Name, Setting.GetType.Name, ex.Message))
                    End Try

                ElseIf mbr.MemberType = MemberTypes.Field Then
                    fe = mbr
                    Try
                        o = fe.GetValue(Setting)

                        If o IsNot Nothing Then

                            If Options And SerializerAdapterFlags.MatchUsePropName Then
                                pr = MatchFieldToProperty(fe.Name, mp)
                            End If

                            ReDim Preserve bfmt.Children(x)

                            If pr IsNot Nothing Then
                                bfmt.Children(x).Name = pr.Name
                                pr = Nothing
                            Else
                                bfmt.Children(x).Name = fe.Name
                            End If

                            SerializeObject(o, bfmt.Children(x))
                            x += 1
                        End If
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try

                End If
            Next

            bfmt.ChildCount = x

        End Sub

        ''' <summary>
        ''' Internally deserialize an object from the binary structure.
        ''' </summary>
        ''' <param name="Setting"></param>
        ''' <param name="bfmt"></param>
        ''' <param name="fieldType"></param>
        ''' <remarks></remarks>
        Private Sub DeserializeObject(ByRef Setting As Object, ByRef bfmt As BIN_SER_FMT, Optional fieldType As System.Type = Nothing)

            Dim fe As FieldInfo = Nothing
            Dim pe As PropertyInfo = Nothing
            Dim pr As PropertyInfo = Nothing

            Dim typeString As String = bfmt.TypeString
            Dim atKnown As Boolean = False

            Try
                If bfmt.DataType = BlobTypes.String Then
                    Setting = bfmt.ToString
                    Return
                ElseIf bfmt.DataType = BlobTypes.Blob Then
                    Setting = bfmt.Data
                    Return
                End If

                If fieldType Is Nothing Then
                    fieldType = FindKnownTypeLike(_KnownTypes, typeString)

                    If fieldType Is Nothing Then
                        If Setting IsNot Nothing Then
                            fieldType = Setting.GetType
                        ElseIf typeString IsNot Nothing Then
                            Try
                                fieldType = System.Type.GetType(typeString, False)
                            Catch ex As Exception
                                Return
                            End Try
                        End If
                    Else
                        atKnown = True
                    End If
                End If

                If Setting Is Nothing And fieldType IsNot Nothing Then

                    If bfmt.DataType = BlobTypes.Image Then
                        Setting = BlobConverter.BytesToImage(bfmt.Data)
                        Return
                    Else
                        If fieldType.IsArray Then
                            Setting = Array.CreateInstance(fieldType.GetElementType, bfmt.ChildCount)
                        Else
                            If atKnown Then
                                Setting = GetInstance(fieldType)
                            Else
                                Setting = GetInstance(fieldType, _KnownTypes)
                            End If
                        End If
                    End If

                    If Setting Is Nothing Then
                        Dim ie As New InstanceHelperEventArgs(fieldType)

                        If bfmt.Data IsNot Nothing AndAlso bfmt.Data.Count > 0 Then
                            ie.SubElementCount = bfmt.Data.Count
                        Else
                            ie.SubElementCount = bfmt.ChildCount
                        End If

                        RaiseEvent InstanceHelper(Me, ie)
                        Setting = ie.Instance
                    End If

                    If Setting IsNot Nothing AndAlso bfmt.Data IsNot Nothing Then
                        Dim isso As ISelfSerializingObject = If(Setting.GetType.GetInterface("ISelfSerializingObject") IsNot Nothing, Setting, Nothing)

                        If isso IsNot Nothing Then
                            isso.Deserialize(bfmt.Data)
                            Return
                        End If
                    End If

                End If

                If fieldType Is Nothing Then Return

                If (bfmt.DataType And BlobTypes.Invalid) = 0 Then
                    Setting = bfmt.Value
                    Return
                End If

                If Setting.GetType.GetInterface("ICollection") IsNot Nothing OrElse Setting.GetType.GetInterface("ICollection`1") IsNot Nothing Then

                    Dim x As Integer = 0,
                    c As Integer = bfmt.ChildCount - 1
                    Dim a As Boolean = Setting.GetType.IsArray

                    Dim arrSetting = Setting
                    Dim subSetting As Object = Nothing

                    'If Not a Then
                    '    a.GetType.GetConstructor(BindingFlags.Public Or BindingFlags.Instance, Nothing, {}, {}).Invoke(Nothing)
                    'End If

                    For x = 0 To c

                        Try
                            subSetting = Nothing
                            DeserializeObject(subSetting, bfmt.Children(x))

                            If subSetting IsNot Nothing Then
                                If a Then
                                    arrSetting(x) = subSetting
                                Else
                                    Setting.Add(subSetting)
                                End If
                            End If
                        Catch ex As Exception
                            MsgBox("Element " & x & ": " & ex.Message)
                        End Try
                    Next

                    Return
                End If

                Dim mp() As PropertyInfo = Nothing
                Dim mi() As MemberInfo = MakeSerializeMap(_Options, fieldType, mp)

                If mi Is Nothing Then Return

                For Each mbr In mi

                    If mbr.MemberType = MemberTypes.Property Then
                        pe = mbr
                        Try
                            Dim o As Object = Nothing

                            If bfmt.HasChild(pe.Name) Then

                                If pe.GetType = GetType(Byte()) Then
                                    pe.SetValue(Setting, bfmt.Data, Nothing)
                                    Continue For
                                End If

                                o = pe.GetValue(Setting, Nothing)
                                DeserializeObject(o, bfmt.ChildByName(pe.Name), pe.PropertyType)

                                If o IsNot Nothing AndAlso TypeOf o Is String Then
                                    o = CType(o, String).Trim(ChrW(0))
                                End If

                                pe.SetValue(Setting, o, Nothing)
                            End If
                        Catch ex As Exception
                            MsgBox("Computer says no: " & ex.Message)
                        End Try

                    ElseIf mbr.MemberType = MemberTypes.Field Then
                        fe = mbr
                        Dim nn As String
                        If Options And SerializerAdapterFlags.MatchUsePropName Then
                            pr = MatchFieldToProperty(fe.Name, mp)
                        End If
                        If pr IsNot Nothing Then
                            nn = pr.Name
                            pr = Nothing
                        Else
                            nn = fe.Name
                        End If

                        Dim o As Object = Nothing

                        If Not bfmt.HasChild(nn) Then Continue For

                        o = fe.GetValue(Setting)
                        DeserializeObject(o, bfmt.ChildByName(nn), fe.FieldType)
                        If TypeOf o Is Char() AndAlso fe.FieldType = GetType(String) Then
                            fe.SetValue(Setting, CStr(o))
                        Else
                            fe.SetValue(Setting, o)
                        End If

                    End If
                Next

            Catch ex As Exception

                MsgBox("I have no idea but, it's " & ex.Message & vbCrLf & "At Property: " & pe.Name)
                Return

            End Try

        End Sub

        ''' <summary>
        ''' Gets the serialization type implemented by the class.  Text = 0, Binary = 1, Xml = 2
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Type As SerializationTypes Implements ISerializerAdapter.Type
            Get
                Return SerializationTypes.Binary
            End Get
        End Property

        Public Property KnownTypes As Type() Implements ISerializerAdapter.KnownTypes
    End Class

#End Region

End Namespace