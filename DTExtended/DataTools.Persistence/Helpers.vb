'' Helper Classes
'' Copyright (C) 2013-2015 Nathaniel N. Moschkin
'' All Rights Reserved

Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Json
Imports System.Reflection

Imports DataTools.Strings
Namespace Persistence

    Public Class Helpers

        Public Shared Function GetAllImplementationsOf(iface As System.Type, assy As Assembly) As System.Type()

            Dim t() As System.Type = assy.GetTypes()

            Dim ot As New List(Of System.Type)
            Dim ichk As System.Type

            '' pre-enumerate all interface and descendents

            Dim ifaces As New List(Of System.Type)

            ifaces.Add(iface)

            For Each chkt In t

                If chkt.IsInterface Then
                    If Not ifaces.Contains(chkt) Then
                        ifaces.Add(chkt)

                        If chkt.BaseType Is Nothing OrElse chkt.BaseType = GetType(Object) Then Continue For

                        If chkt.BaseType <> chkt Then
                            If Not ifaces.Contains(chkt.BaseType) Then ifaces.Add(chkt.BaseType)
                        End If

                    End If

                End If
            Next

            For Each chk In t
                If chk.IsInterface Then Continue For

                For Each chkt In ifaces
                    ichk = chk.GetInterface(chkt.Name)
                    If ichk IsNot Nothing AndAlso ot.Contains(chk) = False Then ot.Add(chk)
                Next

            Next

            Return ot.ToArray

        End Function

        ''' <summary>
        ''' Makes a graph of all DataContract/DataMember member types.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetKnownTypeList(graphType As System.Type, _
                                                Optional ByRef seenTypes As List(Of System.Type) = Nothing, _
                                                Optional flags As SerializerAdapterFlags = SerializerAdapterFlags.DataContract Or _
                                                SerializerAdapterFlags.MatchReadProperty Or _
                                                SerializerAdapterFlags.MatchPublic Or _
                                                SerializerAdapterFlags.NoFields) As System.Type()

            Dim mi() As MemberInfo = Nothing
            Dim pi() As PropertyInfo = Nothing

            Dim p As PropertyInfo

            Dim et As System.Type = Nothing

            mi = MakeSerializeMap(flags, graphType, pi)

            If mi Is Nothing Then Return {}

            Dim l As List(Of System.Type)
            Dim newAdds As New List(Of System.Type)

            If seenTypes Is Nothing Then
                seenTypes = New List(Of System.Type)
            End If

            l = seenTypes

            If l.Contains(graphType) = False Then l.Add(graphType)

            For Each m In mi
                If m.MemberType = MemberTypes.Property Then
                    p = CType(m, PropertyInfo)

                    If seenTypes.Contains(p.PropertyType) Then Continue For

                    If (p.PropertyType <> GetType(String)) _
                        AndAlso (Not p.PropertyType.IsPrimitive) Then

                        l.Add(p.PropertyType)
                        newAdds.Add(p.PropertyType)

                        et = Nothing

                        If IsArray(p.PropertyType, et) Then
                            If et Is Nothing Then Continue For
                            If (et <> GetType(String)) _
                                AndAlso (Not et.IsPrimitive) Then

                                If seenTypes.Contains(et) Then Continue For

                                newAdds.Add(et)
                                l.Add(et)
                            End If
                        End If

                    End If


                End If
            Next

            If l.Count = 0 Then Return {}

            For Each et In newAdds
                If et = graphType Then Continue For
                If (et = GetType(String)) _
                    OrElse (et.IsPrimitive) _
                    OrElse (et.IsEnum) OrElse (IsArray(et)) Then Continue For

                If et.IsClass Then
                    GetKnownTypeList(et, l, flags)
                End If
            Next

            Dim tees() As System.Type

            newAdds.Clear()

            For Each et In l
                If et.IsInterface Then
                    tees = (GetAllImplementationsOf(et, graphType.Assembly))
                    For Each chkt In tees
                        If Not l.Contains(chkt) Then
                            newAdds.Add(chkt)
                        End If
                    Next

                End If
            Next
            l.AddRange(newAdds.ToArray)

            Return l.ToArray

        End Function

        ''' <summary>
        ''' Determines if a type is an array, collection, or list.
        ''' </summary>
        ''' <param name="o"></param>
        ''' <param name="ElementType"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsArray(o As System.Type, Optional ByRef ElementType As System.Type = Nothing) As Boolean

            Dim t As System.Type = o
            Dim ifs As System.Type()
            Dim a As Boolean = False


            If o.IsGenericType Then
                Return IsGenericList(o, ElementType)
            ElseIf o.BaseType IsNot Nothing AndAlso o.BaseType.IsGenericType Then
                Return IsGenericList(o.BaseType, ElementType)
            End If

            ifs = o.GetInterfaces

            For Each iface In ifs

                If (iface.Namespace = "System.Collections.Generic") Then
                    If (iface.GenericTypeArguments.Count = 1) Then
                        Return IsGenericList(iface, ElementType)
                    End If
                    a = True
                End If

                If iface = GetType(ICollection) Then
                    If (iface.GenericTypeArguments.Count = 1) Then
                        Return IsGenericList(o, ElementType)
                    End If
                    a = True
                ElseIf iface = GetType(IList) Then
                    If (iface.GenericTypeArguments.Count = 1) Then
                        Return IsGenericList(o, ElementType)
                    End If
                    a = True
                ElseIf iface = GetType(IEnumerable) Then
                    If (iface.GenericTypeArguments.Count = 1) Then
                        Return IsGenericList(o, ElementType)
                    End If
                    a = True
                End If
            Next

            If a Then
                ElementType = o.GetElementType
                Return True
            End If

            Return False
        End Function

        Public Shared Function IsGenericList(o As System.Type, Optional ByRef ElementType As System.Type = Nothing) As Boolean

            Dim t As System.Type = o
            Dim u As System.Type
            Dim a As Boolean = False

            If t.IsGenericType Then

                Dim ifs() As System.Type = t.GetInterfaces
                u = t.GenericTypeArguments(0)

                For Each iface In ifs
                    If iface = GetType(ICollection) Then
                        a = True
                    ElseIf iface = GetType(IEnumerable) Then
                        a = True
                    ElseIf iface = GetType(IList) Then
                        a = True
                    End If
                Next

                If a = True Then
                    ElementType = u
                    Return True
                End If

            End If

            Return False

        End Function


    End Class



End Namespace

