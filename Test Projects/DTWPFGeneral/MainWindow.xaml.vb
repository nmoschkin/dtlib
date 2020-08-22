Imports System
Imports System.Text
Imports System.IO

Imports DataTools.Strings
Imports DataTools.Interop
Imports DataTools.Memory

Imports DataTools.UnitConversion


Class MainWindow


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        SetValue(ConverterPropertyKey, New ConverterViewModel)


        ' Add any initialization after the InitializeComponent() call.

        Dim mm As New MemPtr(128)

        mm.ZeroMemory()
        mm.ByteAt(16) = 253
        Dim b As Byte = mm.ByteAt(16)

        mm.ByteAt(120) = 33

        mm.ByteAt(0) = 8
        b = mm.ByteAt(0)

        mm.UShortAt(32) = 65535

        mm.Length()
        Dim d() As Long
        d = mm
        mm.Free()

        mm = d

        Erase d
        
        Dim g As Guid = Guid.NewGuid

        mm.GuidAt(1) = g

        Dim s() As Single = mm

        Dim bx() As Byte = mm

        mm.Free()

        mm = bx

        Dim ul() As UInteger = mm

        mm.Free()

        mm = s

        Erase ul
        ul = mm

        mm.Free()

        mm = ul

        Dim g2 As Guid = Guid.Empty
        g2 = mm.GuidAt(1)

        mm.Free()

    End Sub


    Public ReadOnly Property Converter As ConverterViewModel
        Get
            Return GetValue(MainWindow.ConverterProperty)
        End Get
    End Property

    Private Shared ReadOnly ConverterPropertyKey As DependencyPropertyKey = _
                            DependencyProperty.RegisterReadOnly("Converter", _
                            GetType(ConverterViewModel), GetType(MainWindow), _
                            New PropertyMetadata(Nothing))

    Public Shared ReadOnly ConverterProperty As DependencyProperty = _
                           ConverterPropertyKey.DependencyProperty

    Private Sub DoConvert_Click(sender As Object, e As RoutedEventArgs) Handles DoConvert.Click

        Converter.DoConvert()

    End Sub


End Class



Public Class ConverterViewModel
    Inherits DependencyObject

    Private convert As New MetricTool
    Private parseInfo As MetricTool.MetricInfo()

    Public ReadOnly Property ConversionResult As String
        Get
            Return GetValue(ConverterViewModel.ConversionResultProperty)
        End Get
    End Property

    Private Shared ReadOnly ConversionResultPropertyKey As DependencyPropertyKey = _
                            DependencyProperty.RegisterReadOnly("ConversionResult", _
                            GetType(String), GetType(ConverterViewModel), _
                            New PropertyMetadata(Nothing))

    Public Shared ReadOnly ConversionResultProperty As DependencyProperty = _
                           ConversionResultPropertyKey.DependencyProperty

    Public Sub DoConvert()
        convert.Convert(ConversionQuery, Nothing, parseInfo, False)
        SetValue(ConversionResultPropertyKey, convert.Query)
    End Sub

    Public Property ConversionQuery As String
        Get
            Return GetValue(ConversionQueryProperty)
        End Get

        Set(ByVal value As String)
            SetValue(ConversionQueryProperty, value)
        End Set
    End Property

    Public Shared ReadOnly ConversionQueryProperty As DependencyProperty = _
                           DependencyProperty.Register("ConversionQuery", _
                           GetType(String), GetType(ConverterViewModel), _
                           New PropertyMetadata(Nothing))

    Sub New()

    End Sub

End Class