
Imports System.Reflection
Imports DataTools.Memory
Imports DataTools.Memory.Internal


Namespace Persistence

    Public Class ExpGraph

        Public Property KnownTypes As Type()

        <ThreadStatic>
        Private Shared bl As New Blob With {.InBufferMode = True, .StringCatNoNull = True}

        Private _Options As SerializerAdapterFlags =
            SerializerAdapterFlags.DataContract Or
            SerializerAdapterFlags.DataMember Or
            SerializerAdapterFlags.MatchPublic Or
            SerializerAdapterFlags.MatchReadProperty Or
            SerializerAdapterFlags.MatchWriteProperty Or
            SerializerAdapterFlags.NoFields


        Private _RootObject As Object

        Public Property TypeCode As Long
        Public Property Name As String

        Public ReadOnly Property RootObject
            Get
                Return _RootObject
            End Get
        End Property

        Public Property Options As SerializerAdapterFlags
            Get
                Return _Options
            End Get
            Set(value As SerializerAdapterFlags)
                _Options = value
                KnownTypes = Helpers.GetKnownTypeList(_RootObject.GetType, , _Options)
            End Set
        End Property

        Public Sub New(rootObject As Object)

            _RootObject = rootObject
            KnownTypes = Helpers.GetKnownTypeList(_RootObject.GetType, , _Options)

        End Sub

        Private Function GetTypeCharacteristics(t As System.Type) As TypeCharacteristics




        End Function


        Private Function WriteArray(arrObj As Object(), bl As Blob)

            Dim mp As PropertyInfo()
            Dim sgt As Type = arrObj.GetType.GetElementType

            Dim mi = Cache.FindMembers(sgt)

            If mi Is Nothing Then
                mi = MakeSerializeMap(_Options, sgt, mp)
                If mi IsNot Nothing Then Cache.Add(sgt, _Options, mi)
            End If




        End Function

    End Class

    Public Structure TypeCharacteristics

        Public TypeInfo As System.Type

        Public IsClass As Boolean
        Public IsArray As Boolean
        Public IsCollection As Boolean

        Public HasChildren As Boolean

        Public CanFromString As Boolean
        Public CanToString As Boolean

        Public WillNeedCreate As Boolean
        Public WillNeedInstance As Boolean

    End Structure

End Namespace