
Option Explicit On
Option Strict Off
Option Compare Binary

Imports System
Imports System.Text
Imports System.Math
Imports System.ComponentModel
Imports System.Numerics

Namespace UnitConversion

    Public Class FriendlySizeConverter
        Inherits TypeConverter

        Protected Shared _conv As New UnitConversion.MetricTool

        Public Overrides Function CanConvertFrom(context As ITypeDescriptorContext, sourceType As Type) As Boolean
            If sourceType = GetType(String) Then Return True

            Return MyBase.CanConvertFrom(context, sourceType)
        End Function

        Public Overrides Function CanConvertTo(context As ITypeDescriptorContext, destinationType As Type) As Boolean
            If destinationType = GetType(String) Then Return True

            Return MyBase.CanConvertTo(context, destinationType)
        End Function

        Public Overrides Function ConvertFrom(context As ITypeDescriptorContext, culture As Globalization.CultureInfo, value As Object) As Object
            If value.GetType() = GetType(String) Then
                Return CLng(_conv.Convert(value, "bytes"))
            End If

            Return MyBase.ConvertFrom(context, culture, value)
        End Function

        Public Overrides Function ConvertTo(context As ITypeDescriptorContext, culture As Globalization.CultureInfo, value As Object, destinationType As Type) As Object
            If destinationType = GetType(String) Then
                Return PrintFriendlySize(value)
            End If

            Return MyBase.ConvertTo(context, culture, value, destinationType)
        End Function

    End Class

End Namespace

