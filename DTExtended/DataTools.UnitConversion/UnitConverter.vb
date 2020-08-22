
Option Explicit On
Option Strict Off
Option Compare Binary

Imports System
Imports System.Text
Imports System.Math
Imports System.ComponentModel
Imports System.Numerics
Imports DataTools.ExtendedMath

#Region "Unit Conversion"

Namespace UnitConversion

    ''' <summary>
    ''' General, all-purpose unit conversion class.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class MetricTool

        ''' <summary>
        ''' Unit value type converter.
        ''' </summary>
        ''' <remarks></remarks>
        Public Class UnitValueTypeConverter
            Inherits DecimalConverter

            Public Overrides Function CanConvertFrom(context As System.ComponentModel.ITypeDescriptorContext, sourceType As System.Type) As Boolean

                Select Case sourceType
                    Case GetType(String)
                        Return True

                End Select

                If sourceType.IsPrimitive Then Return True

                Return MyBase.CanConvertFrom(context, sourceType)
            End Function

            Public Overrides Function CanConvertTo(context As System.ComponentModel.ITypeDescriptorContext, t As System.Type) As Boolean

                Select Case t
                    Case GetType(String)
                        Return True

                End Select

                If t.IsPrimitive Then Return True

                Return MyBase.CanConvertTo(context, t)
            End Function

            Public Overrides Function ConvertFrom(context As System.ComponentModel.ITypeDescriptorContext, culture As System.Globalization.CultureInfo, value As Object) As Object

                Dim uv As New UnitValue

                If value.GetType() = GetType(String) Then
                    uv.Parse(value)
                    Return uv
                ElseIf IsNumeric(value) Then
                    uv.Value = CType(value, Double)
                    uv.Unit = "px"
                    Return uv
                End If

                Return MyBase.ConvertFrom(context, culture, value)
            End Function

            Public Overrides Function ConvertTo(context As System.ComponentModel.ITypeDescriptorContext, culture As System.Globalization.CultureInfo, value As Object, destinationType As System.Type) As Object

                Dim uv As UnitValue = value, _
                    dt As Type = destinationType

                If destinationType = GetType(String) Then
                    Return value.ToString()
                ElseIf destinationType.IsPrimitive Then
                    Return CType(uv.Value, Double)
                End If

                Return MyBase.ConvertTo(context, culture, value, destinationType)
            End Function

        End Class

        ''' <summary>
        ''' Represents a Value in a given unit or compound unit.
        ''' </summary>
        ''' <remarks></remarks>
        <TypeConverter(GetType(UnitValueTypeConverter))> _
        Public Structure UnitValue

            ''' <summary>
            ''' The value of the measurement.
            ''' </summary>
            ''' <remarks></remarks>
            Public Value As Double

            ''' <summary>
            ''' The string representation of the unit or compound unit of the measurement.
            ''' </summary>
            ''' <remarks></remarks>
            Public Unit As String

            ''' <summary>
            ''' Represents the internal metric info.
            ''' </summary>
            ''' <remarks></remarks>
            Private info As MetricTool.MetricInfo


            ''' <summary>
            ''' Parse the contents structure into a formal unit/value string.
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Overrides Function ToString() As String
                MetricTool.Parse(Me.Value & Me.Unit, info)

                Me.Unit = info.ShortName
                Me.Value = info.Value

                Return info.ShortFormat
            End Function

            ''' <summary>
            ''' Parse a new value.
            ''' </summary>
            ''' <param name="value"></param>
            ''' <remarks></remarks>
            Public Sub Parse(value As String)
                MetricTool.Parse(value, Me.info, "", True)
                Me.Unit = info.ShortName
                Me.Value = info.Value
            End Sub

            Public Function GetDetails() As MetricInfo
                Return info
            End Function

#Region "String CType Operators"

            Public Shared Widening Operator CType(operand As UnitValue) As String
                Return operand.ToString
            End Operator

            Public Shared Narrowing Operator CType(operand As String) As UnitValue
                Dim u As New UnitValue
                u.Parse(operand)

                Return u
            End Operator

#End Region

#Region "Numeric CType Operators"

            Public Shared Function NumberToUnitValue(operand As Double) As UnitValue
                Dim u As New UnitValue
                MetricTool.Parse(operand & "px", u.info)

                u.Unit = u.info.ShortName
                u.Value = u.info.Value

                Return u
            End Function

            Public Shared Widening Operator CType(operand As MetricTool.UnitValue) As Long
                Return CInt(operand.Value)
            End Operator

            Public Shared Narrowing Operator CType(operand As Long) As MetricTool.UnitValue
                Return NumberToUnitValue(CDbl(operand))
            End Operator

            Public Shared Widening Operator CType(operand As MetricTool.UnitValue) As Integer
                Return CInt(operand.Value)
            End Operator

            Public Shared Narrowing Operator CType(operand As Integer) As MetricTool.UnitValue
                Return NumberToUnitValue(CDbl(operand))
            End Operator

            Public Shared Widening Operator CType(operand As MetricTool.UnitValue) As Short
                Return CInt(operand.Value)
            End Operator

            Public Shared Narrowing Operator CType(operand As Short) As MetricTool.UnitValue
                Return NumberToUnitValue(CDbl(operand))
            End Operator

            Public Shared Widening Operator CType(operand As MetricTool.UnitValue) As Double
                Return CInt(operand.Value)
            End Operator

            Public Shared Narrowing Operator CType(operand As Double) As MetricTool.UnitValue
                Return NumberToUnitValue(CDbl(operand))
            End Operator

            'Public Shared Widening Operator CType(operand As MetricTool.UnitValue) As Double
            '    Return CInt(operand.Value)
            'End Operator

            'Public Shared Narrowing Operator CType(operand As Double) As MetricTool.UnitValue
            '    Return NumberToUnitValue(CDbl(operand))
            'End Operator

#End Region

#Region "Equality Operators"

            Public Overrides Function Equals(obj As Object) As Boolean
                Dim u As UnitValue = CType(obj, UnitValue)

                If u.Value <> Value Then Return False
                If u.Unit <> Unit Then Return False

                Return True
            End Function

            Public Shared Operator =(operand1 As UnitValue, operand2 As UnitValue) As Boolean
                Return operand1.Equals(operand2)
            End Operator

            Public Shared Operator <>(operand1 As UnitValue, operand2 As UnitValue) As Boolean
                Return Not operand1.Equals(operand2)
            End Operator

#End Region

        End Structure

        <TypeConverter(GetType(ExpandableObjectConverter))> _
        Public Class MetricUnit
            Implements ICloneable
            Private _Measures As String = ""

            Public Property Measures() As String
                Get
                    Return _Measures
                End Get
                Set(value As String)
                    _Measures = TitleCase(value)
                End Set
            End Property

            Private _Name As String = ""

            Public Property Name() As String
                Get
                    Return _Name
                End Get
                Set(value As String)
                    _Name = TitleCase(value)
                End Set
            End Property

            Private _Prefix As String = ""

            Public Property Prefix() As String
                Get
                    Return _Prefix
                End Get
                Set(value As String)
                    _Prefix = value
                End Set
            End Property

            Private _PluralName As String = ""

            Public Property PluralName() As String
                Get
                    Return _PluralName
                End Get
                Set(value As String)
                    _PluralName = TitleCase(value)
                End Set
            End Property

            '' Indicates this is a base unit to which all other units in this category convert
            Private _IsBase As Boolean = False

            Public Property IsBase() As Boolean
                Get
                    Return _IsBase
                End Get
                Set(value As Boolean)
                    _IsBase = value
                End Set
            End Property

            '' for derived bases.  $ is used to denote a variable.
            Private _Modifies As String = ""

            Public Property Modifies() As String
                Get
                    Return _Modifies
                End Get
                Set(value As String)
                    _Modifies = TitleCase(value)
                End Set
            End Property

            Private _Multiplier As Double = 0.0#

            Public Property Multiplier() As Double
                Get
                    Return _Multiplier
                End Get
                Set(value As Double)
                    _Multiplier = value
                End Set
            End Property

            Private _Offset As Double = 0.0#

            Public Property Offset() As Double
                Get
                    Return _Offset
                End Get
                Set(value As Double)
                    _Offset = value
                End Set
            End Property

            Private _OffsetFirst As Boolean = False

            Public Property OffsetFirst() As Boolean
                Get
                    Return _OffsetFirst
                End Get
                Set(value As Boolean)
                    _OffsetFirst = value
                End Set
            End Property

            '' Set to True for derived bases with equations.

            Private _Derived As Boolean = False

            Public Property Derived() As Boolean
                Get
                    Return _Derived
                End Get
                Set(value As Boolean)
                    _Derived = value
                End Set
            End Property

            '' Equation for derived bases.  $ is used to denote a variable.

            Private _Equation As String = ""

            Public Property Equation() As String
                Get
                    Return _Equation
                End Get
                Set(value As String)
                    _Equation = value
                End Set
            End Property

            Public Function Clone() As Object Implements System.ICloneable.Clone
                Return Me.MemberwiseClone
            End Function

            Public Overrides Function ToString() As String
                Return Me.PluralName & " (" & Me.Measures & ")"
            End Function

        End Class

        <TypeConverter(GetType(ExpandableObjectConverter))> _
        Public Class MetricInfo
            Implements ICloneable

            Private _Value As Double

            Public Property Value As Double
                Get
                    Return _Value
                End Get
                Set(value As Double)
                    _Value = value
                End Set
            End Property

            Private _BaseValue As Double

            Public Property BaseValue As Double
                Get
                    Return _BaseValue
                End Get
                Set(value As Double)
                    _BaseValue = value
                End Set
            End Property

            Private _Multiplier As Double

            Public Property Multiplier As Double
                Get
                    Return _Multiplier
                End Get
                Set(value As Double)
                    _Multiplier = value
                End Set
            End Property

            Private _BaseUnit As String

            Public Property BaseUnit As String
                Get
                    Return _BaseUnit
                End Get
                Set(value As String)
                    _BaseUnit = value
                End Set
            End Property

            Private _Name As String

            Public Property Name As String
                Get
                    Return _Name
                End Get
                Set(value As String)
                    _Name = value
                End Set
            End Property

            Private _PluralName As String

            Public Property PluralName As String
                Get
                    Return _PluralName
                End Get
                Set(value As String)
                    _PluralName = value
                End Set
            End Property

            Private _ShortName As String

            Public Property ShortName As String
                Get
                    Return _ShortName
                End Get
                Set(value As String)
                    _ShortName = value
                End Set
            End Property

            Private _Measures As String

            Public Property Measures As String
                Get
                    Return _Measures
                End Get
                Set(value As String)
                    _Measures = value
                End Set
            End Property

            Private _Format As String

            Public Property Format As String
                Get
                    Return _Format
                End Get
                Set(value As String)
                    _Format = value
                End Set
            End Property

            Private _ShortFormat As String

            Public Property ShortFormat As String
                Get
                    Return _ShortFormat
                End Get
                Set(value As String)
                    _ShortFormat = value
                End Set
            End Property

            Private _Unit As New MetricUnit

            Public Property Unit As MetricUnit
                Get
                    Return _Unit
                End Get
                Set(value As MetricUnit)
                    _Unit = value
                End Set
            End Property

            Public Overrides Function ToString() As String
                Return _Unit.ToString & " " & _Value
            End Function

            Public Function Clone() As Object Implements System.ICloneable.Clone
                Dim objNew As MetricInfo = Me.MemberwiseClone
                objNew.Unit = _Unit.Clone

                Return objNew
            End Function

        End Class

        Private Shared _RoundingDigits As Integer = 4

        Private _Info As New MetricInfo
        Private _Convert As String
        Private _LongFormat As Boolean = True

        Private _ConvIn() As MetricInfo
        Private _ConvOut() As MetricInfo

        Private MyUnits As UnitCollection

        Private Shared _Units As New UnitCollection()

        Public Shared Prefixes() As String = {"", "deci", "centi", "milli", "micro", "nano", "pico", "femto", "atto", "zepto", "yocto", "deca", "hecto", "kilo", "mega", "giga", "tera", "peta", "exa", "zetta", "yotta", "kibi", "mebi", "gibi", "tebi", "pebi", "exbi", "zebi", "yobi"}
        Public Shared ShortPrefixes() As String = {"", "d", "c", "m", "μ", "n", "p", "f", "a", "z", "y", "da", "h", "k", "M", "G", "T", "P", "E", "Z", "Y", "Ki", "Mi", "Gi", "Ti", "Pi", "Ei", "Zi", "Yi"}
        Public Shared Multipliers() As Double = {10 ^ 0, 10 ^ -1, 10 ^ -2, 10 ^ -3, 10 ^ -6, 10 ^ -9, 10 ^ -12, 10 ^ -15, 10 ^ -18, 10 ^ -21, 10 ^ -24, 10 ^ 1, 10 ^ 2, 10 ^ 3, 10 ^ 6, 10 ^ 9, 10 ^ 12, 10 ^ 15, 10 ^ 18, 10 ^ 21, 10 ^ 24, 2 ^ 10, 2 ^ 20, 2 ^ 30, 2 ^ 40, 2 ^ 50, 2 ^ 60, 2 ^ 70, 2 ^ 80}

        <Browsable(True)> _
        Public Shared ReadOnly Property Units As UnitCollection
            Get
                Return _Units
            End Get
        End Property

        <Browsable(True)> _
        Public Property ConversionQuery As MetricInfo()
            Get
                Return _ConvIn
            End Get
            Set(value As MetricInfo())
                _ConvIn = value
            End Set
        End Property

        <Browsable(True)> _
        Public Property ConversionResult As MetricInfo()
            Get
                Return _ConvOut
            End Get
            Set(value As MetricInfo())
                _ConvOut = value
            End Set
        End Property

        <Browsable(True)> _
        Public Property LongFormat As Boolean
            Get
                Return _LongFormat
            End Get
            Set(value As Boolean)
                _LongFormat = value
                Convert(Query)
                Format = Format
            End Set
        End Property

        <Browsable(True)> _
        Public Shared Property RoundingDigits As Integer
            Get
                Return _RoundingDigits
            End Get
            Set(value As Integer)
                _RoundingDigits = value
            End Set
        End Property

        <Browsable(True)> _
        Public Property AvailableUnits As UnitCollection
            Get
                If MyUnits Is Nothing Then Return _Units
                Return MyUnits
            End Get
            Set(value As UnitCollection)
                MyUnits = value
            End Set
        End Property

        <Browsable(True)> _
        Public Property Format As String
            Get
                If _LongFormat = True Then
                    Return _Info.Format
                Else
                    Return _Info.ShortFormat
                End If
            End Get
            Set(value As String)
                Parse(value, _Info)
            End Set
        End Property

        Public Function ApplyEquation(u As MetricUnit, ByRef ErrorText As String, ParamArray vars() As Double) As MetricInfo

            Dim i As Integer, _
                c As Integer

            Dim j As Integer

            Dim s As String = u.Equation, _
                v As String

            Dim objInfo As MetricInfo = Nothing

            c = vars.Length

            For i = 0 To c
                v = "$" & (i + 1)
                j = s.IndexOf(v)

                If j = -1 Then
                    ErrorText = "Missing variable " & v & " in equation, or too many parameters passed."
                    Return Nothing
                End If

                Do Until j = -1
                    s = s.Replace(v, "" & vars(i))
                Loop
            Next

            j = s.IndexOf("$")

            If j <> -1 Then
                ErrorText = "Too few parameters passed."
                Return Nothing
            End If

            Dim p As New MathExpressionParser

            If p.ParseOperations(s) = False Then
                ErrorText = p.ErrorText
                Return Nothing
            End If

            s = "" & p.Value & u.Name
            Parse(s, objInfo)

            Return objInfo

        End Function

        ''' <summary>
        ''' Discerns whether or not a value is an x per y (or x/y) value and returns the parsed contents.
        ''' </summary>
        ''' <param name="value">Value string to analyze.</param>
        ''' <param name="MustMeasure">Imperitive measurement units array.  The string must match this many units of these exact types to parse correctly.</param>
        ''' <returns>Parsed MetricInfo array.</returns>
        ''' <remarks></remarks>
        Public Function IsPer(value As String, Optional MustMeasure() As MetricInfo = Nothing, Optional parseMath As Boolean = True) As MetricInfo()

            Dim s() As String, _
                i As Integer = 0

            Dim m() As MetricInfo = Nothing
            Dim boo As Boolean = False

            Dim perScan As String()

            If parseMath Then
                perScan = {"per"}
            Else
                perScan = {"per", "/"}
            End If

            For Each pp In perScan

                If value.IndexOf(pp) <> -1 Then

                    s = BatchParse(value, pp)
                    ReDim m(s.Length - 1)

                    For i = 0 To s.Length - 1
                        m(i) = New MetricInfo
                        s(i) = Trim(s(i))
                        If i <> 0 Then If FVal(s(i)) = 0 Then s(i) = "1" & s(i)

                        If MustMeasure Is Nothing Then
                            Parse(s(i), m(i))
                        Else
                            If i <= MustMeasure.Length - 1 Then _
                                Parse(s(i), m(i), MustMeasure(i).Measures)
                        End If

                    Next
                    boo = True

                End If

            Next

            If Not boo Then
                m = {New MetricInfo}

                If MustMeasure Is Nothing Then
                    Parse(value, m(0))
                Else
                    If i <= MustMeasure.Length - 1 Then _
                        Parse(value, m(0), MustMeasure(i).Measures)
                End If
            End If

            Return m

        End Function

        Public Function ComparePer(m1() As MetricInfo, m2() As MetricInfo) As Boolean
            Dim i As Integer

            If m1.Length <> m2.Length Then Return False

            For i = 0 To m1.Length - 1
                If m1(i).Measures <> m2(i).Measures Then
                    _Convert = "Cannot convert between " & SeparateCamel(m1(i).Measures) & " and " & SeparateCamel(m2(i).Measures) & "."
                    Return False

                End If
            Next

            Return True
        End Function

        ''' <summary>
        ''' Get or set the conversion query.
        ''' </summary>
        ''' <returns></returns>
        Public Property Query As String
            Get
                Return _Convert
            End Get
            Set(value As String)
                Convert(value)
            End Set
        End Property

        ''' <summary>
        ''' Returns the value in the current unit.
        ''' </summary>
        ''' <returns></returns>
        <Browsable(True)> _
        Public ReadOnly Property Value As Double
            Get
                Return _Info.Value
            End Get
        End Property

        ''' <summary>
        ''' Returns the <see cref="MetricInfo"/> object.
        ''' </summary>
        ''' <returns></returns>
        <Browsable(True)> _
        Public Property Info As MetricInfo
            Get
                Return _Info
            End Get
            Set(value As MetricInfo)
                _Info = value
            End Set
        End Property

        Public Overloads Function Convert(inputValue As MetricInfo, outputType As String, ByRef outputValue As MetricInfo) As Double
            Dim s As String = inputValue.BaseValue & inputValue.BaseUnit & "=" & outputType
            Convert(s)
            outputValue = _ConvOut(0).Clone
            Return _ConvOut(0).Value
        End Function

        Public Overloads Function Convert(inputValue As MetricInfo, outputType As MetricUnit, ByRef outputValue As MetricInfo) As Double
            Dim s As String = inputValue.BaseValue & inputValue.BaseUnit & "=" & outputType.Name
            Convert(s)
            outputValue = _ConvOut(0).Clone
            Return _ConvOut(0).Value
        End Function

        Public Overloads Function Convert(inputValue As MetricInfo, outputType As MetricUnit) As Double
            Dim s As String = inputValue.BaseValue & inputValue.BaseUnit & "=" & outputType.Name
            Convert(s)
            Return _ConvOut(0).Value
        End Function

        Public Overloads Function Convert(inputValue As Double, outputType As MetricUnit) As Double
            Dim s As String = inputValue & "=" & outputType.Name
            Convert(s)
            Return _ConvOut(0).Value
        End Function

        Public Overloads Function Convert(inputValue As String, outputType As String) As Double

            Dim s As String

            s = inputValue & "=" & outputType
            Convert(s)

            Return _ConvOut(0).Value

        End Function

        Public Overloads Function Convert(inputValue As Double, inputType As String, Optional outputType As String = "px") As Double

            Dim s As String

            s = "" & inputValue & inputType & "=" & outputType
            Convert(s)

            Return _ConvOut(0).Value

        End Function

        ''' <summary>
        ''' Convert a series of unit representations into another form.
        ''' For exmaple: 1 liter = gallons - or - 100 km/h = mi/h.  The answers will appear in the <see cref="MetricTool.Query">Query</see> variable.
        ''' </summary>
        ''' <param name="value">The string to parse.</param>
        ''' <param name="InputInfo">The parsed detailed input value/unit array.</param>
        ''' <param name="OutputInfo">The detailed output value/unit array.</param>
        ''' <param name="parseMath">Specifies that the value string contains a mathematical equation.</param>
        ''' <remarks></remarks>
        Public Overloads Sub Convert(value As String, Optional ByRef InputInfo() As MetricInfo = Nothing, Optional ByRef OutputInfo() As MetricInfo = Nothing, Optional parseMath As Boolean = True)

            Dim s() As String = Nothing, _
                    i As Integer, _
                    v As Double = 0.0#, _
                    vv() As Double

            Dim i1() As MetricInfo, _
                i2() As MetricInfo
            If value = Nothing Then Return

            i = value.ToLower.IndexOf("convert ")
            If i <> -1 Then
                value = Trim(value.Substring(i + 8))
            End If

            If value.IndexOf("=") = -1 Then
                value = value.Replace(" to ", "=")
                value = value.Replace(" TO ", "=")
                value = value.Replace(" To ", "=")
                value = value.Replace(" is ", "=")
                value = value.Replace(" IS ", "=")
                value = value.Replace(" Is ", "=")
                value = value.Replace(" and ", "=")
                value = value.Replace(" AND ", "=")
                value = value.Replace(" And ", "=")
                value = value.Replace(" what ", "=")
                value = value.Replace(" WHAT ", "=")
                value = value.Replace(" What ", "=")
                value = value.Replace(" how many ", "=")
                value = value.Replace(" how much ", "=")
                value = value.Replace(" HOW MANY ", "=")
                value = value.Replace(" HOW MUCH ", "=")
                value = value.Replace(" How Many ", "=")
                value = value.Replace(" How Much ", "=")
                value = value.Replace("to to", "=")
                value = value.Replace("to to", "=")
            End If

            value = OneSpace(value)

            If value.IndexOf("=") <> -1 Then
                s = BatchParse(value, "=")
            ElseIf value.IndexOf(",") <> -1 Then
                s = BatchParse(value, ",")
            End If

            If s Is Nothing Then Return

            If s.Length <> 2 Then Return

            s(0) = Trim(s(0))
            If Not IsNumeric(s(1)) Then s(1) = "1 " & Trim(s(1))

            i1 = IsPer(s(0), , parseMath)
            i2 = IsPer(s(1), i1, parseMath)

            If ComparePer(i1, i2) = False Then Return

            'If i2.Length >= 2 Then
            '    For i = 1 To i2.Length - 1
            '        Parse("" & i1(i).Value & i2(i).PluralName, i2(i))
            '    Next
            'End If

            If i1.Length = 1 Then
                v = i1(0).BaseValue

                With i2(0).Unit
                    If .Modifies = i1(0).BaseUnit OrElse .Name = i1(0).BaseUnit Then

                        If .Name = i1(0).BaseUnit Then
                            v = i1(0).BaseValue
                        Else
                            If .OffsetFirst Then
                                v /= .Multiplier
                                v -= .Offset
                            Else
                                v -= .Offset
                                v /= .Multiplier
                            End If


                        End If

                    End If

                    v /= i2(0).Multiplier

                End With
                v = Math.Round(v, _RoundingDigits)

                Parse(v & " " & i2(0).Name, i2(0), i1(0).Measures, parseMath)

                If _LongFormat Then
                    _Convert = "" & i1(0).Format & " is " & i2(0).Format
                Else
                    _Convert = "" & i1(0).ShortFormat & " is " & i2(0).ShortFormat
                End If

            Else
                ReDim vv(i1.Length - 1)

                For i = 0 To i1.Length - 1
                    vv(i) = i1(i).BaseValue

                    With i2(i).Unit
                        If .Modifies = i1(i).BaseUnit OrElse .Name = i1(i).BaseUnit Then

                            If .Name = i1(i).BaseUnit Then
                                vv(i) = i1(i).BaseValue

                            Else
                                If .OffsetFirst Then
                                    vv(i) /= .Multiplier
                                    vv(i) -= .Offset
                                Else
                                    vv(i) -= .Offset
                                    vv(i) /= .Multiplier
                                End If
                            End If
                        End If

                        vv(i) /= i2(i).Multiplier

                    End With

                    If Math.Round(vv(i), _RoundingDigits) <> 0 Then _
                        vv(i) = Math.Round(vv(i), _RoundingDigits)

                    If i = 0 Then
                        v = vv(i)
                    Else
                        'i2(i).Value = 1
                        'If (i1(i).Value <> 0) Then vv(i) /= i1(i).Value
                        v /= vv(i)
                    End If

                Next

                v = Math.Round(v, _RoundingDigits)

                _Convert = i1(0).Value.ToString("#,##0.####") & " "

                For i = 0 To i1.Length - 1

                    If _LongFormat Then
                        If (i <> 0) Then _Convert &= " per " & i1(i).Value.ToString("#,##0.####") & " "
                        If (i1(i).Value <> 1) Then
                            _Convert &= i1(i).PluralName
                        Else
                            _Convert &= i1(i).Name
                        End If
                    Else
                        If (i <> 0) Then _Convert &= "/"
                        _Convert &= i1(i).ShortName

                    End If
                Next

                _Convert &= " is "

                _Convert &= v.ToString("#,##0.####") & " "

                i2(0).Value = v

                For i = 0 To i2.Length - 1

                    If _LongFormat Then
                        If (i <> 0) Then _Convert &= " per " & i2(i).Value.ToString("#,##0.####") & " "

                        If (i2(i).Value <> 1) Then
                            _Convert &= i2(i).PluralName
                        Else
                            _Convert &= i2(i).Name
                        End If
                    Else
                        If (i <> 0) Then _Convert &= "/"
                        _Convert &= i2(i).ShortName
                    End If
                Next


            End If

            ConversionQuery = i1
            ConversionResult = i2

            InputInfo = i1.Clone
            OutputInfo = i2.Clone

        End Sub

        Public Shared Sub WordsTest(Example As String)
            Dim s() As String

            s = Words(Example)
            MsgBox("There are " & s.Length & " words, starting with " & s(0))



        End Sub

        Public Shared Function GetMultiplier(prefix As String) As Double
            Dim i As Integer, _
                c As Integer

            c = Prefixes.Length - 1

            If prefix.Length <= 2 Then
                For i = 0 To c
                    If prefix = ShortPrefixes(i) Then
                        Return Multipliers(i)
                    End If
                Next
            End If

            prefix = prefix.ToLower

            For i = 0 To c

                If prefix = Prefixes(i) Then
                    Return Multipliers(i)
                End If
            Next

            ' it's always safe to return 1
            Return 1

        End Function

        Public Shared Function Parse(ByVal value As String, ByRef info As MetricInfo, Optional ByVal MustMeasure As String = "", Optional ByVal parseMath As Boolean = True) As Double
            Dim mMath As MathExpressionParser = Nothing

            Dim ua() As MetricUnit
            Dim markFail As Boolean = False

            Dim ch() As Char, _
                i As Integer, _
                c As Integer

            Dim ss() As String, _
                x As Integer, _
                y As Integer

            Dim s As String = "", _
                cb As String = "", _
                r As Double = 0.0#

            Dim t As String

            Dim n1 As Integer = 0, _
                n2 As Integer = 0

            If ParseMath Then mMath = New MathExpressionParser
            If MustMeasure IsNot Nothing Then MustMeasure = MustMeasure.ToLower

            If info Is Nothing OrElse info.Name Is Nothing Then
                info = New MetricInfo

                If String.IsNullOrEmpty(MustMeasure) = False Then
                    Parse(value, info, , parseMath)
                End If
            End If

            value = Trim(value)
            ch = value

            c = ch.Length - 1
            If value = "" Then Return 0

            If ParseMath = False Then

                Do Until IsNumeric(ch(c))
                    c -= 1
                    If c < 0 Then Exit Do
                Loop
            Else
                Do Until IsNumeric(ch(c)) Or (ch(c) = ")")
                    c -= 1
                    If c < 0 Then Exit Do
                Loop
            End If

            If c < 0 Then Return Nothing

            s = Trim(value.Substring(c + 1))
            value = value.Substring(0, c + 1)

            i = 0

            If ParseMath Then
                mMath.ParseOperations(value)
                r = mMath.Value
            Else
                If IsHex(value, i) Then
                    r = i
                Else
                    r = FVal(value)
                End If
            End If

            If s.Length > 1 Then _
                If s.Substring(s.Length - 1) = "." Then s = s.Substring(0, s.Length - 1)

            r = Math.Round(r, _RoundingDigits)
            t = s.ToLower

            With info
                .Value = r

                If MustMeasure <> "" Then
                    ua = GetUnitsArray(MustMeasure)
                Else
                    ua = Units.InnerArray
                End If

                If ua Is Nothing Then Return Double.NaN

                For n1 = 0 To ua.Length - 1
                    Erase ss

                    If ua(n1).Prefix = Nothing Then
                        ReDim ss(0)
                        ss(0) = ""
                    Else
                        ss = BatchParse(ua(n1).Prefix, ",")
                    End If

                    For n2 = 0 To ShortPrefixes.Length - 1

                        y = ss.Length - 1

                        For x = 0 To y
                            ss(x) = Trim(ss(x))
                            cb = ShortPrefixes(n2) & ss(x)
                            If s = cb Or s = cb & "s" Then
                                If (MustMeasure = "") Or (MustMeasure = ua(n1).Measures.ToLower) Then

                                    .ShortName = ShortPrefixes(n2) & ss(0)
                                    .Name = Prefixes(n2) & ua(n1).Name

                                    Exit For
                                End If
                                ' we found it!
                            End If

                        Next

                        If x <= y Then Exit For
                    Next
                    If n2 < ShortPrefixes.Length Then Exit For

                    For n2 = 0 To Prefixes.Length - 1

                        cb = Prefixes(n2) & ua(n1).Name.ToLower()
                        If t = cb Or t = cb & "s" Then
                            If (MustMeasure = "") Or (MustMeasure = ua(n1).Measures.ToLower) Then
                                .Name = cb
                                .ShortName = ShortPrefixes(n2) & ss(0)
                                Exit For
                                ' we found it!
                            End If

                        End If

                        cb = Prefixes(n2) & ua(n1).PluralName.ToLower()
                        If t = cb Or t = cb & "s" Then
                            If (MustMeasure = "") Or (MustMeasure = ua(n1).Measures.ToLower) Then
                                .PluralName = cb
                                .ShortName = ShortPrefixes(n2) & ss(0)
                                .Name = Prefixes(n2) & ua(n1).Name.ToLower()
                                Exit For
                                ' we found it!
                            End If
                        End If


                    Next

                    If n2 < Prefixes.Length Then Exit For
                Next

                If (n1 < Units.Count) And (n2 < Prefixes.Length) Then
                    .Unit = ua(n1)
                    .Multiplier = Multipliers(n2)
                    .BaseValue = .Value * .Multiplier
                    .BaseUnit = ua(n1).Name
                    .Measures = ua(n1).Measures

                    .Format = "" & .Value.ToString("#,##0.##") & " "
                    .Name = TitleCase(.Name)

                    If (.Value <> 1) And (ua(n1).PluralName <> "") Then
                        .Format &= TitleCase(Prefixes(n2) & ua(n1).PluralName.ToLower())
                    Else
                        .Format &= TitleCase(.Name)
                    End If

                    .ShortFormat = "" & .Value.ToString("#,##0.##") & " " & .ShortName

                    If ua(n1).Modifies <> Nothing Then
                        If ua(n1).OffsetFirst = True Then
                            .BaseValue += ua(n1).Offset
                            .BaseValue *= ua(n1).Multiplier
                        Else
                            .BaseValue *= ua(n1).Multiplier
                            .BaseValue += ua(n1).Offset
                        End If
                        .BaseUnit = ua(n1).Modifies
                    End If

                    If (ua(n1).PluralName <> "") Then
                        If Prefixes(n2) = "" Then
                            .PluralName = TitleCase(ua(n1).PluralName)
                        Else
                            .PluralName = TitleCase(Prefixes(n2) & ua(n1).PluralName.ToLower())
                        End If
                    End If

                End If

                If markFail Then Return Double.NaN

                Return .Value
            End With

        End Function

        Sub New(Optional ByVal WithOwnUnits As Boolean = False)
            If WithOwnUnits Then MyUnits = GetUnits()

            With _Info
                .BaseValue = 0
                .Multiplier = 1
                .Value = 0
            End With
        End Sub

        Shared Sub New()
            Dim objUnit As MetricUnit

            objUnit = AddUnit()
            With objUnit
                .Name = "liter"
                .PluralName = "liters"
                .Prefix = "L,litre,litres,ltr"
                .Measures = "volume"
                .IsBase = True
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "gallon"
                .PluralName = "gallons"
                .Prefix = "gl,gal"
                .Multiplier = 3.785411784
                .Modifies = "liter"
                .Measures = "volume"
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "cup"
                .PluralName = "cups"
                .Prefix = "cup"
                .Multiplier = 0.2365882365
                .Modifies = "liter"
                .Measures = "volume"
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "tablespoon"
                .PluralName = "tablespoons"
                .Prefix = "tbsp"
                .Multiplier = 0.01478676478125
                .Modifies = "liter"
                .Measures = "volume"
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "teaspoon"
                .PluralName = "teaspoons"
                .Prefix = "tsp"
                .Multiplier = 0.00492892159375
                .Modifies = "liter"
                .Measures = "volume"
            End With


            objUnit = AddUnit()
            With objUnit
                .Name = "meter squared"
                .PluralName = "meters squared"
                .Prefix = "msq"
                .Measures = "area"
                .IsBase = True
            End With


            objUnit = AddUnit()
            With objUnit
                .Name = "hectare"
                .PluralName = "hectares"
                .Prefix = "ha"
                .Multiplier = 10000
                .Modifies = "meter squared"
                .Measures = "area"
            End With


            objUnit = AddUnit()
            With objUnit
                .Name = "acre"
                .PluralName = "acres"
                .Prefix = "ac"
                .Multiplier = 4046.8564224
                .Modifies = "meter squared"
                .Measures = "area"
            End With


            objUnit = AddUnit()
            With objUnit
                .Name = "square foot"
                .PluralName = "square feet"
                .Prefix = "sqft,sft,ftsq"
                .Multiplier = 0.09290304
                .Modifies = "meter squared"
                .Measures = "area"
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "square inch"
                .PluralName = "square inches"
                .Prefix = "insq,sqin"
                .Multiplier = 0.00064516
                .Modifies = "meter squared"
                .Measures = "area"
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "square mile"
                .PluralName = "square miles"
                .Prefix = "misq,sqmi,sqm"
                .Multiplier = 2589.9
                .Modifies = "meter squared"
                .Measures = "area"
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "quart"
                .PluralName = "quarts"
                .Prefix = "qt"
                .Multiplier = 0.946352946
                .Modifies = "liter"
                .Measures = "volume"
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "pint"
                .PluralName = "pints"
                .Prefix = "p"
                .Multiplier = 0.473176473
                .Modifies = "liter"
                .Measures = "volume"
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "meter"
                .PluralName = "meters"
                .Prefix = "m"
                .Measures = "length"
                .IsBase = True
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "gram"
                .PluralName = "grams"
                .Prefix = "#,##0.####"
                .Measures = "mass"
                .IsBase = True
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "pound"
                .PluralName = "pounds"
                .Modifies = "gram"
                .Offset = 0
                .Multiplier = 453.59237
                .Prefix = "lb"
                .Measures = "mass"
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "stone"
                .PluralName = "stone"
                .Modifies = "gram"
                .Offset = 0
                .Multiplier = 6350
                .Prefix = "stone"
                .Measures = "mass"
            End With


            objUnit = AddUnit()
            With objUnit
                .Name = "metric ton"
                .PluralName = "metric tons"
                .Modifies = "gram"
                .Offset = 0
                .Multiplier = 1000 * 1000
                .Prefix = "mtn"
                .Measures = "mass"
            End With


            objUnit = AddUnit()
            With objUnit
                .Name = "ton"
                .PluralName = "tons"
                .Modifies = "gram"
                .Offset = 0
                .Multiplier = 907 * 1000
                .Prefix = "tn"
                .Measures = "mass"
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "week"
                .PluralName = "weeks"
                .Modifies = "second"
                .Offset = 0
                .Multiplier = 60 * 60 * 24 * 7
                .Prefix = "wk"
                .Measures = "time"
            End With


            objUnit = AddUnit()
            With objUnit
                .Name = "ounce"
                .PluralName = "ounces"
                .Modifies = "gram"
                .Offset = 0
                .Multiplier = 28.349523125
                .Prefix = "oz"
                .Measures = "mass"
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "second"
                .PluralName = "seconds"
                .Prefix = "s,sec"
                .Measures = "time"
                .IsBase = True
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "ampere"
                .PluralName = "amperes"
                .Prefix = "A"
                .Measures = "electric current"
                .IsBase = True
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "kelvin"
                .PluralName = ""
                .Prefix = "K"
                .Measures = "thermodynamic temperature"
                .IsBase = True
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "mole"
                .PluralName = "moles"
                .Prefix = "mol"
                .Measures = "amount of substance"
                .IsBase = True
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "candela"
                .PluralName = ""
                .Prefix = "cd"
                .Measures = "luminous intensity"
                .IsBase = True
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "byte"
                .PluralName = "bytes"
                .Modifies = "bit"
                .Multiplier = 8
                .Prefix = "B"
                .Measures = "binary data"
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "bit"
                .PluralName = "bits"
                .Prefix = "b"
                .Measures = "binary data"
                .IsBase = True
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "parsec"
                .PluralName = "parsecs"
                .Prefix = "pc"
                .Measures = "length"
                .Modifies = "meter"
                .Multiplier = (30.857 * (10 ^ 12)) * 1000
                .Offset = 0
                .OffsetFirst = False
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "light year"
                .PluralName = "light years"
                .Prefix = "ly"
                .Measures = "length"
                .Modifies = "meter"
                .Multiplier = 9.4607 * (10 ^ 15)
                .Offset = 0
                .OffsetFirst = False
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "celsius"
                .PluralName = ""
                .Prefix = "C"
                .Measures = "thermodynamic temperature"
                .Modifies = "kelvin"
                .Multiplier = 1
                .Offset = 273.15
                .OffsetFirst = False
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "fahrenheit"
                .PluralName = ""
                .Prefix = "F"
                .Measures = "thermodynamic temperature"
                .Modifies = "kelvin"
                .Multiplier = (5 / 9)
                .Offset = 459.67
                .OffsetFirst = True
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "foot"
                .PluralName = "feet"
                .Prefix = "ft,'"
                .Measures = "length"
                .Modifies = "meter"
                .Multiplier = 0.3048
                .Offset = 0
                .OffsetFirst = False
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "inch"
                .PluralName = "inches"
                .Prefix = "in," & Chr(34)
                .Measures = "length"
                .Modifies = "meter"
                .Multiplier = 0.0254
                .Offset = 0
                .OffsetFirst = False
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "pixel"
                .PluralName = "pixels"
                .Prefix = "px"
                .Measures = "length"
                .Modifies = "meter"
                .Multiplier = 0.00026458333333333336
                .Offset = 0
                .OffsetFirst = False
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "point"
                .PluralName = "points"
                .Prefix = "pt"
                .Measures = "length"
                .Modifies = "meter"
                .Multiplier = 0.00035278
                .Offset = 0
                .OffsetFirst = False
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "pica"
                .PluralName = "picas"
                .Prefix = "P"
                .Measures = "length"
                .Modifies = "meter"
                .Multiplier = 0.004233
                .Offset = 0
                .OffsetFirst = False
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "mile"
                .PluralName = "miles"
                .Prefix = "mi"
                .Measures = "length"
                .Modifies = "meter"
                .Multiplier = 1609.344
                .Offset = 0
                .OffsetFirst = False
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "yard"
                .PluralName = "yards"
                .Prefix = "yd"
                .Measures = "length"
                .Modifies = "meter"
                .Multiplier = 0.9144
                .Offset = 0
                .OffsetFirst = False
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "minute"
                .PluralName = "minutes"
                .Prefix = "min"
                .Measures = "time"
                .Modifies = "second"
                .Multiplier = 60
                .Offset = 0
                .OffsetFirst = False
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "hour"
                .PluralName = "hours"
                .Prefix = "h"
                .Measures = "time"
                .Modifies = "second"
                .Multiplier = 60 * 60
                .Offset = 0
                .OffsetFirst = False
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "day"
                .PluralName = "days"
                .Prefix = "d"
                .Measures = "time"
                .Modifies = "second"
                .Multiplier = 60 * 60 * 24
                .Offset = 0
                .OffsetFirst = False
            End With


            objUnit = AddUnit()
            With objUnit
                .Name = "year"
                .PluralName = "years"
                .Prefix = "y"
                .Measures = "time"
                .Modifies = "second"
                .Multiplier = 60 * 60 * 24 * 365.25
                .Offset = 0
                .OffsetFirst = False
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "annus"
                .PluralName = "annum"
                .Prefix = "a"
                .Measures = "time"
                .Modifies = "second"
                .Multiplier = 60 * 60 * 24 * 365.25
                .Offset = 0
                .OffsetFirst = False
            End With


            objUnit = AddUnit()
            With objUnit
                .Name = "century"
                .PluralName = "centuries"
                .Prefix = "C."
                .Measures = "time"
                .Modifies = "second"
                .Multiplier = 60 * 60 * 24 * 365.25 * 100
                .Offset = 25 * 60 * 60 * 24
                .OffsetFirst = False
            End With


            objUnit = AddUnit()
            With objUnit
                .Name = "millennium"
                .PluralName = "millennia"
                .Prefix = "M."
                .Measures = "time"
                .Modifies = "second"
                .Multiplier = 60 * 60 * 24 * 365.25 * 1000
                .Offset = 0.0#
                .OffsetFirst = False
            End With

            objUnit = AddUnit()
            With objUnit
                .Name = "aeon"
                .PluralName = "aeons"
                .Prefix = "AE."
                .Measures = "time"
                .Modifies = "second"
                .Multiplier = (60 * 60 * 24 * 365.25 * 1000) * 1000000D
                .Offset = 0.0#
                .OffsetFirst = False
            End With

            SortCategories(True)

        End Sub

        Private Shared Sub SortCategories(Optional ByVal BaseUnitsFirst As Boolean = False)
            '' Sorts all categories alphabetically, optionally putting all base units first

            Dim c1 As New UnitCollection, _
                objUnit As MetricUnit

            Dim s1() As String, _
                s2() As String

            Dim i As Integer, _
                c As Integer

            Dim j As Integer, _
                d As Integer

            s1 = GetCategories()
            c = s1.Length - 1

            '' each category starts with its base unit, and then modifying units come next alphabetically, by category alphabetically
            If BaseUnitsFirst = False Then
                For i = 0 To c
                    s2 = GetUnitNames(s1(i))

                    d = s2.Length - 1
                    For j = 0 To d
                        objUnit = FindUnit(s2(j), s1(i))
                        If objUnit.IsBase = True Then
                            c1.Add(objUnit)
                            Exit For
                        End If
                    Next

                    d = s2.Length - 1
                    For j = 0 To d
                        objUnit = FindUnit(s2(j), s1(i))
                        If objUnit.IsBase = False Then _
                            c1.Add(objUnit)
                    Next

                Next

                _Units = c1
                Return
            Else
                '' this is the alternative way, listing all base units by category alphabetically, first, then all other units by category alphabetically
                For i = 0 To c
                    s2 = GetUnitNames(s1(i))

                    d = s2.Length - 1
                    For j = 0 To d
                        objUnit = FindUnit(s2(j), s1(i))
                        If objUnit.IsBase = True Then
                            c1.Add(objUnit)
                            Exit For
                        End If
                    Next

                Next

                For i = 0 To c
                    s2 = GetUnitNames(s1(i))

                    d = s2.Length - 1
                    For j = 0 To d
                        objUnit = FindUnit(s2(j), s1(i))
                        If objUnit.IsBase = False Then _
                            c1.Add(objUnit)

                    Next

                Next

                _Units = c1
                Return

            End If
        End Sub

        Public Shared Function AddUnit(Optional Measures As String = "", Optional Name As String = "", Optional PluralName As String = "", Optional Prefix As String = "", Optional Modifies As String = "", Optional Multiplier As Double = 0.0#, Optional Offset As Double = 0.0#, Optional OffsetFirst As Boolean = False, Optional IsBase As Boolean = False) As MetricUnit

            Dim b As New MetricUnit

            With b
                .Measures = Measures
                .IsBase = IsBase
                .Name = Name
                .PluralName = PluralName
                .Prefix = Prefix
                .Modifies = Modifies
                .Multiplier = Multiplier
                .Offset = Offset
                .OffsetFirst = OffsetFirst
            End With

            _Units.Add(b)
            Return b

        End Function

        Public Shared Function GetCategories() As String()
            Dim s() As String = Nothing, _
                u As MetricUnit

            Dim c As Integer = -1

            For Each u In _Units
                If c = -1 Then
                    ReDim s(0)
                    s(0) = TitleCase(u.Measures)
                    c = 1
                Else
                    If s.Contains(TitleCase(u.Measures)) = False Then
                        ReDim Preserve s(c)
                        s(c) = TitleCase(u.Measures)
                        c += 1
                    End If
                End If
            Next

            Array.Sort(s)
            Return s

        End Function

        <Description("Get all unit names for a category.")> _
        Public Shared Function GetUnitNames(Category As String, Optional ByVal ExcludeBaseUnit As Boolean = False) As String()
            Dim c() As String = Nothing, _
                n As Integer = 0
            Dim u As MetricUnit

            For Each u In _Units
                If (u.Measures.ToLower = Category.ToLower) And ((u.IsBase = False) Or (ExcludeBaseUnit = False)) Then
                    ReDim Preserve c(n)
                    c(n) = TitleCase(u.Name)
                    n += 1
                End If
            Next

            Array.Sort(c)
            Return c
        End Function

        <Description("Get all base unit names.")> _
        Public Shared Function GetBaseUnitNames() As String()
            Dim c() As String = Nothing, _
                n As Integer = 0
            Dim u As MetricUnit

            For Each u In _Units
                If u.IsBase = True Then
                    ReDim Preserve c(n)
                    c(n) = TitleCase(u.Name)
                    n += 1
                End If
            Next

            Array.Sort(c)
            Return c
        End Function

        Public Shared Function HasCategory(Category As String) As Boolean
            Dim u As MetricUnit

            For Each u In _Units
                If u.Measures.ToLower = Category.ToLower Then Return True
            Next

            Return False

        End Function

        <Description("Get all base units for all categories.")> _
        Public Shared Function GetBaseUnits() As UnitCollection
            Dim c As New UnitCollection
            Dim u As MetricUnit

            For Each u In _Units
                If u.IsBase = True Then _
                    c.Add(u.Clone)
            Next

            Return c
        End Function

        <Description("Get all units for a category.")> _
        Public Shared Function GetUnits(Optional Category As String = "") As UnitCollection
            Dim uc As New UnitCollection
            Dim a() As MetricUnit, _
                i As Integer, _
                c As Integer

            a = _Units.InnerArray
            c = a.Length - 1
            Category = TitleCase(Category)

            For i = 0 To c
                If a(i).Measures = Category Or Category = "" Then
                    uc.Add(a(i).Clone)
                End If
            Next

            Return uc
        End Function

        <Description("Get all units for a category as an array.")> _
        Public Shared Function GetUnitsArray(Optional Category As String = "") As MetricUnit()
            Dim a() As MetricUnit, _
                i As Integer, _
                c As Integer

            Dim b() As MetricUnit = Nothing, _
                n As Integer = 0

            a = _Units.InnerArray
            c = a.Length - 1

            Category = Category.ToLower

            For i = 0 To c
                If NoSpace(a(i).Measures.ToLower) = Category Or Category = "" Then
                    ReDim Preserve b(n)
                    b(n) = a(i).Clone
                    n += 1
                End If
            Next

            Return b
        End Function

        <Description("Get the base unit for a specific category.")> _
        Public Shared Function GetBaseUnit(Category As String) As MetricUnit
            Dim u As MetricUnit

            For Each u In _Units
                If u.IsBase = True And (u.Measures.ToLower = Category.ToLower) Then _
                    Return (u.Clone)
            Next

            Return Nothing
        End Function

        <Description("Find an exact unit.")> _
        Public Shared Function FindUnit(Unit As String, Optional ByVal MustMeasure As String = "") As MetricUnit
            Dim u As MetricUnit
            Dim s() As String
            Dim i As Integer, _
                c As Integer

            For Each u In _Units
                If (u.Measures.ToLower = MustMeasure.ToLower) Or (MustMeasure = "") Then
                    If u.Name.ToLower = Unit.ToLower Then Return u.Clone
                    If u.PluralName.ToLower = Unit.ToLower Then Return u.Clone
                    s = BatchParse(u.Prefix, ",")
                    c = s.Length - 1
                    For i = 0 To c
                        If s(i) = Unit Then Return u.Clone
                    Next
                End If
            Next

            Return Nothing
        End Function

    End Class

    <DefaultProperty("Item")> _
    Public Class UnitCollection
        Inherits CollectionBase

        Public Class UnitCollectionEventArgs
            Inherits EventArgs

            Private _Item As MetricTool.MetricUnit = Nothing
            Private _Cancel As Boolean = False

            Public Property Cancel As Boolean
                Get
                    Return _Cancel
                End Get
                Set(value As Boolean)
                    _Cancel = value
                End Set
            End Property

            Public Sub New(item As MetricTool.MetricUnit)
                _Item = item
            End Sub

            Public ReadOnly Property Item() As MetricTool.MetricUnit
                Get
                    Return _Item
                End Get
            End Property

        End Class

        Protected _Parent As MetricTool

        Public Event ObjectAdd(sender As Object, e As UnitCollectionEventArgs)
        Public Event ObjectRemove(sender As Object, e As UnitCollectionEventArgs)
        Public Event ClearItems(sender As Object, e As UnitCollectionEventArgs)

        <Browsable(False)> _
        Public ReadOnly Property InnerArray As MetricTool.MetricUnit()
            Get
                Return InnerList.ToArray(GetType(MetricTool.MetricUnit))
            End Get
        End Property

        <Browsable(False)> _
        Public ReadOnly Property Parent As MetricTool
            Get
                Return _Parent
            End Get
        End Property

        Default Public Property Item(index As Integer) As MetricTool.MetricUnit
            Get
                On Error Resume Next
                Return CType(List(index), MetricTool.MetricUnit)
            End Get
            Set(value As MetricTool.MetricUnit)
                List(index) = value
            End Set
        End Property

        Public Shadows Sub Clear()
            Dim e As UnitCollectionEventArgs
            e = New UnitCollectionEventArgs(Nothing)

            RaiseEvent ClearItems(Me, e)
            If e.Cancel = True Then Return

            List.Clear()
        End Sub

        Public Function Add(value As MetricTool.MetricUnit) As Integer
            Dim i As Integer
            Dim e As New UnitCollectionEventArgs(value)

            i = List.Add(value)

            RaiseEvent ObjectAdd(Me, e)
            If e.Cancel = True Then
                List.Remove(value)
                i = -1
            End If

            Return i
        End Function 'Add

        Public Function IndexOf(value As MetricTool.MetricUnit) As Integer
            Return List.IndexOf(value)
        End Function 'IndexOf

        Public Sub Insert(index As Integer, value As MetricTool.MetricUnit)

            Dim i As Integer = List.Add(value)
            Dim e As New UnitCollectionEventArgs(value)

            RaiseEvent ObjectAdd(Me, New UnitCollectionEventArgs(value))
            If e.Cancel = True Then Return

            List.Insert(index, value)
        End Sub 'Insert

        Public Sub Remove(value As MetricTool.MetricUnit)
            Dim e As UnitCollectionEventArgs
            e = New UnitCollectionEventArgs(value)

            RaiseEvent ObjectRemove(Me, e)
            If e.Cancel = True Then Return

            List.Remove(value)
        End Sub 'Remove

        Public Function Contains(value As MetricTool.MetricUnit) As Boolean
            Return List.Contains(value)
        End Function

        Protected Overrides Sub OnValidate(value As Object)
            If Not GetType(MetricTool.MetricUnit).IsAssignableFrom(value.GetType()) Then
                Throw New ArgumentException("value must be implement MetricTool.MetricUnit.", "value")
            End If
        End Sub 'OnValidate 

        Public Sub New(parent As MetricTool)
            _Parent = parent
        End Sub

        Public Sub New()

        End Sub

    End Class

End Namespace

#End Region