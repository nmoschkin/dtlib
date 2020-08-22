Imports System.ComponentModel
Imports System.Runtime.Serialization
Imports System.Text
Imports System.Drawing.Design

Namespace Persistence

    <Serializable, DataContract, TypeConverter(GetType(VersionObjectConverter)), Editor(GetType(UITypeEditor), GetType(VersionObjectTypeEditor)), Description("Version information")>
    Public Class VersionInfoObject
        Implements ICloneable, INotifyPropertyChanged, ISelfSerializingObject

        Private _Major As Integer
        Private _Minor As Integer
        Private _Revision As Integer

        <ThreadStatic>
        Private Shared bl As New DataTools.Memory.Blob With {.InBufferMode = True, .StringCatNoNull = True}

        ''' <summary>
        ''' Major version (0-9999)
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DataMember, Browsable(True), Description("Major Version (0-9999)")>
        Public Property Major As Integer
            Get
                Return _Major
            End Get
            Set(value As Integer)
                _Major = value
                Check()
                OnPropertyChanged("Major")
            End Set
        End Property

        ''' <summary>
        ''' Minor version (0-9999)
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DataMember, Browsable(True), Description("Minor Version (0-9999)")>
        Public Property Minor As Integer
            Get
                Return _Minor
            End Get
            Set(value As Integer)
                _Minor = value
                Check()
                OnPropertyChanged("Minor")
            End Set
        End Property

        ''' <summary>
        ''' Revision (0-99999)
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DataMember, Browsable(True), Description("Revision (0-99999)")>
        Public Property Revision As Integer
            Get
                Return _Revision
            End Get
            Set(value As Integer)
                _Revision = value
                Check()
                OnPropertyChanged("Revision")
            End Set
        End Property

        Public Function Clone() As Object Implements ICloneable.Clone
            Return Me.MemberwiseClone
        End Function

        Private Sub Check()
            If _Major < 0 Then _Major = 0
            If _Major > 9999 Then _Major = 9999
            If _Minor < 0 Then _Minor = 0
            If _Minor > 9999 Then _Minor = 9999
            If _Revision < 0 Then _Revision = 0
            If _Revision > 99999 Then _Revision = 99999

        End Sub

        Public Sub IncRev(Optional val As Integer = 1)
            _Revision += val
            Check()
            OnPropertyChanged("Revision")
        End Sub

        Public Sub IncMaj(Optional val As Integer = 1)
            _Major += val
            Check()
            OnPropertyChanged("Major")
        End Sub

        Public Sub IncMin(Optional val As Integer = 1)
            _Minor += val
            Check()
            OnPropertyChanged("Minor")
        End Sub

        Public Sub DateToMinRev()
            Dim d As Date = Now
            Dim s As String
            s = d.Month.ToString("00") & d.Day.ToString("00")
            _Minor = d.Year
            _Revision = Integer.Parse(s)

            OnPropertyChanged("Revision")
            OnPropertyChanged("Minor")
        End Sub

        Public Sub New()
            _Major = 1
            DateToMinRev()
        End Sub

        Public Sub New(maj As Integer, min As Integer, rev As Integer)
            _Major = maj
            _Minor = min
            _Revision = rev
            OnPropertyChanged("Revision")
            OnPropertyChanged("Minor")
            OnPropertyChanged("Major")

        End Sub

        Public Sub New(version As String)
            Dim b() As String = BatchParse(version, ".")

            If b Is Nothing Then Return

            If b.Length >= 1 Then
                Major = Integer.Parse(b(0))
            End If

            If b.Length >= 2 Then
                Minor = Integer.Parse(b(1))
            End If

            If b.Length >= 3 Then
                Revision = Integer.Parse(b(2))
            End If

            OnPropertyChanged("Revision")
            OnPropertyChanged("Minor")
            OnPropertyChanged("Major")
        End Sub

        Public Overrides Function ToString() As String
            Dim sb As New StringBuilder

            sb.Append(_Major.ToString)
            sb.Append(".")
            sb.Append(_Minor.ToString("00"))
            sb.Append(".")
            sb.Append(_Revision.ToString("0000"))

            ToString = sb.ToString
        End Function

        Public Shared Widening Operator CType(operand As String) As VersionInfoObject
            Return New VersionInfoObject(operand)
        End Operator

        Public Shared Narrowing Operator CType(operand As VersionInfoObject) As String
            Return operand.ToString
        End Operator

        Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements INotifyPropertyChanged.PropertyChanged

        Protected Overloads Sub OnPropertyChanged(e As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(e))
        End Sub

        Protected Overloads Sub OnPropertyChanged(e As PropertyChangedEventArgs)
            RaiseEvent PropertyChanged(Me, e)
        End Sub

        Public Function Serialize() As Byte() Implements ISelfSerializingObject.Serialize

            bl.Length = 0
            bl.ClipReset()

            bl &= _Major
            bl &= _Minor
            bl &= _Revision

            Return bl
        End Function

        Public Sub Deserialize(data() As Byte) Implements ISelfSerializingObject.Deserialize

            bl.Length = 0
            bl.ClipReset()
            bl &= data

            _Major = bl.IntegerAt(0)
            _Minor = bl.IntegerAt(1)
            _Revision = bl.IntegerAt(2)

        End Sub

    End Class

    Public Class VersionObjectConverter
        Inherits TypeConverter

        Public Overrides Function CanConvertFrom(context As ITypeDescriptorContext, sourceType As Type) As Boolean
            If sourceType = GetType(String) Then Return True

            Return MyBase.CanConvertFrom(context, sourceType)
        End Function

        Public Overrides Function CanConvertTo(context As ITypeDescriptorContext, destinationType As Type) As Boolean
            If destinationType = GetType(String) Then Return True

            Return MyBase.CanConvertTo(context, destinationType)
        End Function

        Public Overrides Function ConvertFrom(context As ITypeDescriptorContext, culture As Globalization.CultureInfo, value As Object) As Object
            If value.GetType = GetType(String) Then
                Return New VersionInfoObject(value)
            End If

            Return MyBase.ConvertFrom(context, culture, value)
        End Function

        Public Overrides Function ConvertTo(context As ITypeDescriptorContext, culture As Globalization.CultureInfo, value As Object, destinationType As Type) As Object
            If destinationType = GetType(String) Then
                Return CType(value, VersionInfoObject).ToString
            End If

            Return MyBase.ConvertTo(context, culture, value, destinationType)
        End Function

    End Class

End Namespace