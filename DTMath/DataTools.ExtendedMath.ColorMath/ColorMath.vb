'' ColorMath class.  Copyright (C) 1999-2014 Nathaniel Moschkin.
'' This is the 4th writing of this class.

'' This most recent version uses clarified equations harvested from Wikipedia.
'' The remainder of the code was written much prior... and was correctly
'' achieved through much trial and error.

Option Explicit On
Option Strict Off
Option Compare Binary

Imports System.Math
Imports System.Drawing
Imports System.Runtime.InteropServices

Namespace ColorMath

    Public Class ColorMath

        Public Const SF_VALUE = &H0
        Public Const SF_SATURATION = &H1
        Public Const SF_HUE = &H2

        Public Shared Function Cex(ByVal Color As System.Drawing.Color) As String

            Dim s As String = ""
            Dim b() As Byte = BitConverter.GetBytes(Color.ToArgb)

            For i As Integer = 0 To 3
                s &= b.ToString("X2")
            Next

            Return "#" & s.ToLower

        End Function

        Public Shared Function UnCex(value As String) As System.Drawing.Color
            Dim s As String()
            ReDim s(2)

            If value.Substring(0, 1) <> "#" Then Return Color.Transparent
            value = value.Substring(1)

            If value.Length = 2 Then
                s(0) = value.Substring(0, 2)
                s(1) = value.Substring(0, 2)
                s(2) = value.Substring(0, 2)
            Else
                s(0) = value.Substring(0, 2)
                s(1) = value.Substring(2, 2)
                s(2) = value.Substring(4, 2)
            End If

            Dim r As Integer = Val("&H" & s(0))
            Dim g As Integer = Val("&H" & s(1))
            Dim b As Integer = Val("&H" & s(2))

            Return System.Drawing.Color.FromArgb(255, r, g, b)

        End Function

        Public Shared Function UnCexToMediaColor(value As String) As System.Windows.Media.Color
            Dim s As String()
            ReDim s(3)

            If value.Substring(0, 1) = "#" Then value = value.Substring(1)

            If value.Length = 2 Then
                s(0) = "ff"
                s(1) = value.Substring(0, 2)
                s(2) = value.Substring(0, 2)
                s(3) = value.Substring(0, 2)
            ElseIf value.Length = 8 Then
                s(0) = value.Substring(0, 2)
                s(1) = value.Substring(2, 2)
                s(2) = value.Substring(4, 2)
                s(3) = value.Substring(6, 2)
            ElseIf value.Length = 6 Then
                s(0) = "ff"
                s(1) = value.Substring(0, 2)
                s(2) = value.Substring(2, 2)
                s(3) = value.Substring(4, 2)
            Else
                Return Nothing
            End If

            Dim a As Byte = Val("&H" & s(0))
            Dim r As Byte = Val("&H" & s(1))
            Dim g As Byte = Val("&H" & s(2))
            Dim b As Byte = Val("&H" & s(3))

            Return System.Windows.Media.Color.FromArgb(a, r, g, b)

        End Function

        Public Overloads Shared Sub GetRGB(ByVal Color As System.Drawing.Color, ByRef Red As Object, ByRef Green As Object, ByRef Blue As Object)

            Dim i As Integer = &H0&
            Dim crColor As Integer = Color.ToArgb()

            Red = crColor And &HFF
            Green = (crColor >> 8) And &HFF
            Blue = (crColor >> 16) And &HFF

        End Sub

        Public Overloads Shared Sub GetRGB(ByVal Color As Integer, ByRef Red As Object, ByRef Green As Object, ByRef Blue As Object)

            Red = Color And &HFF
            Green = (Color >> 8) And &HFF
            Blue = (Color >> 16) And &HFF

        End Sub

        '' Single Convert ColorRef to RGB

        Public Shared Sub ColorToRGB(ByVal Color As System.Drawing.Color, ByRef Bits As RGBDATA)
            Dim b() As Byte
            b = BitConverter.GetBytes(Color.ToArgb)

            Bits.Red = b(2)
            Bits.Green = b(1)
            Bits.Blue = b(0)

        End Sub

        Public Shared Sub ColorToARGB(ByVal Color As System.Drawing.Color, ByRef Bits As ARGBDATA)
            Dim b() As Byte
            b = BitConverter.GetBytes(Color.ToArgb)

            Bits.Alpha = b(3)
            Bits.Red = b(2)
            Bits.Green = b(1)
            Bits.Blue = b(0)
        End Sub

        '' Single Convert RGB to ColorRef

        Public Shared Function RGBToColor(Bits As RGBDATA) As System.Drawing.Color
            Return System.Drawing.Color.FromArgb(255, Bits.Red, Bits.Green, Bits.Blue)
        End Function

        Public Shared Function ARGBToColor(Bits As ARGBDATA) As System.Drawing.Color
            Return System.Drawing.Color.FromArgb(Bits.Alpha, Bits.Red, Bits.Green, Bits.Blue)
        End Function

        '' Single Convert ColorRef to RGB-reversed

        Public Shared Sub ColorToBGR(ByVal Color As System.Drawing.Color, ByRef Bits As BGRDATA)

            Dim tibs As RGBDATA = New RGBDATA

            ColorToRGB(Color, tibs)

            Bits.Blue = tibs.Blue
            Bits.Red = tibs.Red
            Bits.Green = tibs.Green

        End Sub

        Public Shared Sub ColorToBGRA(ByVal Color As System.Drawing.Color, ByRef Bits As BGRADATA)

            Dim tibs As ARGBDATA = New ARGBDATA

            ColorToARGB(Color, tibs)

            Bits.Alpha = tibs.Blue
            Bits.Blue = tibs.Blue
            Bits.Red = tibs.Red
            Bits.Green = tibs.Green

        End Sub

        '' Single Convert RGB-reversed to ColorRef

        Public Shared Function BGRAToColor(Bits As BGRADATA) As System.Drawing.Color

            Dim tibs As ARGBDATA = New ARGBDATA
            tibs.Alpha = Bits.Alpha
            tibs.Red = Bits.Red
            tibs.Blue = Bits.Blue
            tibs.Green = Bits.Green

            Return ARGBToColor(tibs)

        End Function

        Public Shared Function BGRToColor(Bits As BGRDATA) As System.Drawing.Color

            Dim tibs As RGBDATA = New RGBDATA
            tibs.Red = Bits.Red
            tibs.Blue = Bits.Blue
            tibs.Green = Bits.Green

            Return RGBToColor(tibs)

        End Function

#Region "They removed good functionality!  We're putting it back Max and Min for more than two!"

        Public Overloads Shared Function Min(ParamArray vars() As Double) As Double

            Dim d As Double,
            i As Integer,
            c As Integer = vars.Length - 1

            If c <= 0 Then Return Double.NaN

            d = vars(0)
            For i = 1 To c
                d = System.Math.Min(d, vars(i))
            Next

            Return d
        End Function

        Public Overloads Shared Function Max(ParamArray vars() As Double) As Double

            Dim d As Double,
            i As Integer,
            c As Integer = vars.Length - 1

            If c <= 0 Then Return Double.NaN

            d = vars(0)
            For i = 1 To c
                d = System.Math.Max(d, vars(i))
            Next

            Return d
        End Function

        Public Overloads Shared Function Min(ParamArray vars() As Single) As Single

            Dim d As Single,
            i As Integer,
            c As Integer = vars.Length - 1

            If c <= 0 Then Return Single.NaN

            d = vars(0)
            For i = 1 To c
                d = System.Math.Min(d, vars(i))
            Next

            Return d
        End Function

        Public Overloads Shared Function Max(ParamArray vars() As Single) As Single

            Dim d As Single,
            i As Integer,
            c As Integer = vars.Length - 1

            If c <= 0 Then Return Single.NaN

            d = vars(0)
            For i = 1 To c
                d = System.Math.Max(d, vars(i))
            Next

            Return d
        End Function

        Public Overloads Shared Function Min(ParamArray vars() As Decimal) As Decimal

            Dim d As Decimal,
            i As Integer,
            c As Integer = vars.Length - 1

            If c <= 0 Then Return Decimal.Zero

            d = vars(0)
            For i = 1 To c
                d = System.Math.Min(d, vars(i))
            Next

            Return d
        End Function

        Public Overloads Shared Function Max(ParamArray vars() As Decimal) As Decimal

            Dim d As Decimal,
            i As Integer,
            c As Integer = vars.Length - 1

            If c <= 0 Then Return Decimal.Zero

            d = vars(0)
            For i = 1 To c
                d = System.Math.Max(d, vars(i))
            Next

            Return d
        End Function

        Public Overloads Shared Function Min(ParamArray vars() As Long) As Long

            Dim d As Long,
            i As Long,
            c As Long = vars.Length - 1

            If c <= 0 Then Return 0

            d = vars(0)
            For i = 1 To c
                d = System.Math.Min(d, vars(i))
            Next

            Return d
        End Function

        Public Overloads Shared Function Max(ParamArray vars() As Long) As Long

            Dim d As Long,
            i As Long,
            c As Long = vars.Length - 1

            If c <= 0 Then Return 0

            d = vars(0)
            For i = 1 To c
                d = System.Math.Max(d, vars(i))
            Next

            Return d
        End Function

        Public Overloads Shared Function Min(ParamArray vars() As Integer) As Integer

            Dim d As Integer,
            i As Integer,
            c As Integer = vars.Length - 1

            If c <= 0 Then Return 0

            d = vars(0)
            For i = 1 To c
                d = System.Math.Min(d, vars(i))
            Next

            Return d
        End Function

        Public Overloads Shared Function Max(ParamArray vars() As Integer) As Integer

            Dim d As Integer,
            i As Integer,
            c As Integer = vars.Length - 1

            If c <= 0 Then Return 0

            d = vars(0)
            For i = 1 To c
                d = System.Math.Max(d, vars(i))
            Next

            Return d
        End Function

        Public Overloads Shared Function Min(ParamArray vars() As Short) As Short

            Dim d As Short,
            i As Integer,
            c As Integer = vars.Length - 1

            If c <= 0 Then Return 0

            d = vars(0)
            For i = 1 To c
                d = System.Math.Min(d, vars(i))
            Next

            Return d
        End Function

        Public Overloads Shared Function Max(ParamArray vars() As Short) As Short

            Dim d As Short,
            i As Integer,
            c As Integer = vars.Length - 1

            If c <= 0 Then Return 0

            d = vars(0)
            For i = 1 To c
                d = System.Math.Max(d, vars(i))
            Next

            Return d
        End Function

        Public Overloads Shared Function Min(ParamArray vars() As Byte) As Byte

            Dim d As Byte,
            i As Integer,
            c As Integer = vars.Length - 1

            If c <= 0 Then Return 0

            d = vars(0)
            For i = 1 To c
                d = System.Math.Min(d, vars(i))
            Next

            Return d
        End Function

        Public Overloads Shared Function Max(ParamArray vars() As Byte) As Byte

            Dim d As Byte,
            i As Integer,
            c As Integer = vars.Length - 1

            If c <= 0 Then Return 0

            d = vars(0)
            For i = 1 To c
                d = System.Math.Max(d, vars(i))
            Next

            Return d
        End Function

        Public Overloads Shared Function Min(ParamArray vars() As ULong) As ULong

            Dim d As ULong,
            i As Integer,
            c As Integer = vars.Length - 1

            If c <= 0 Then Return 0

            d = vars(0)
            For i = 1 To c
                d = System.Math.Min(d, vars(i))
            Next

            Return d
        End Function

        Public Overloads Shared Function Max(ParamArray vars() As ULong) As ULong

            Dim d As ULong,
            i As Integer,
            c As Integer = vars.Length - 1

            If c <= 0 Then Return 0

            d = vars(0)
            For i = 1 To c
                d = System.Math.Max(d, vars(i))
            Next

            Return d
        End Function

        Public Overloads Shared Function Min(ParamArray vars() As UInteger) As UInteger

            Dim d As UInteger,
            i As Integer,
            c As Integer = vars.Length - 1

            If c <= 0 Then Return 0

            d = vars(0)
            For i = 1 To c
                d = System.Math.Min(d, vars(i))
            Next

            Return d
        End Function

        Public Overloads Shared Function Max(ParamArray vars() As UInteger) As UInteger

            Dim d As UInteger,
            i As Integer,
            c As Integer = vars.Length - 1

            If c <= 0 Then Return 0

            d = vars(0)
            For i = 1 To c
                d = System.Math.Max(d, vars(i))
            Next

            Return d
        End Function

        Public Overloads Shared Function Min(ParamArray vars() As UShort) As UShort

            Dim d As UShort,
            i As Integer,
            c As Integer = vars.Length - 1

            If c <= 0 Then Return 0

            d = vars(0)
            For i = 1 To c
                d = System.Math.Min(d, vars(i))
            Next

            Return d
        End Function

        Public Overloads Shared Function Max(ParamArray vars() As UShort) As UShort

            Dim d As UShort,
            i As Integer,
            c As Integer = vars.Length - 1

            If c <= 0 Then Return 0

            d = vars(0)
            For i = 1 To c
                d = System.Math.Max(d, vars(i))
            Next

            Return d
        End Function

        Public Overloads Shared Function Min(ParamArray vars() As SByte) As SByte

            Dim d As SByte,
            i As Integer,
            c As Integer = vars.Length - 1

            If c <= 0 Then Return 0

            d = vars(0)
            For i = 1 To c
                d = System.Math.Min(d, vars(i))
            Next

            Return d
        End Function

        Public Overloads Shared Function Max(ParamArray vars() As SByte) As SByte

            Dim d As SByte,
            i As Integer,
            c As Integer = vars.Length - 1

            If c <= 0 Then Return 0

            d = vars(0)
            For i = 1 To c
                d = System.Math.Max(d, vars(i))
            Next

            Return d
        End Function

#End Region

        Private Structure clrstr

            Public a As Single

            Public r As Single

            Public g As Single

            Public b As Single

            Public Sub New(c As System.Drawing.Color)
                a = c.A
                r = c.R
                g = c.G
                b = c.B
            End Sub


            Public Sub New(c As System.Windows.Media.Color)
                a = c.ScA
                r = c.ScR
                g = c.ScG
                b = c.ScB
            End Sub

            Public Sub New(i As Integer)

                r = i And &HFF
                g = (i >> 8) And &HFF
                b = (i >> 16) And &HFF
                a = (i >> 24) And &HFF

            End Sub

        End Structure

        Public Shared Sub ColorToHSV(ByVal color As System.Windows.Media.color, ByRef hsv As HSVDATA)
            ColorToHSV(New clrstr(color), hsv)
        End Sub

        Public Shared Sub ColorToHSV(ByVal color As System.Drawing.Color, ByRef hsv As HSVDATA)
            ColorToHSV(New clrstr(color), hsv)
        End Sub

        Public Shared Sub ColorToHSV(ByVal color As Integer, ByRef hsv As HSVDATA)
            ColorToHSV(New clrstr(color), hsv)
        End Sub


        Private Shared Sub ColorToHSV(ByVal Color As clrstr, ByRef hsv As HSVDATA)

            '' http://en.wikipedia.org/wiki/HSL_and_HSV#Hue_and_chroma
            '' I adapted the equation from the one in Wikipedia.
            '' I wish I could offer a better explanation.  But this isn't wikipedia, and they'd do a better job.

            '' who says you can't divide by zero?
            On Error Resume Next

            Dim hue As Double,
            sat As Double,
            val As Double

            Dim Mn As Double,
            Mx As Double

            Dim r As Double,
            g As Double,
            b As Double

            Dim chroma As Double

            Dim rgb As RGBDATA


            r = CDbl(Color.r)
            g = CDbl(Color.g)
            b = CDbl(Color.b)

            Mn = Min(r, g, b)
            Mx = Max(r, g, b)

            chroma = Mx - Mn
            val = Mx

            If (Mn = Mx) Then

                hsv.Hue = -1

                val = Mn

                With hsv
                    Select Case val

                        Case Is <= 0.5

                            .Saturation = 1
                            .Value = (510 * val) / 360

                        Case Is <= 1

                            val = 1 - val

                            .Value = 1
                            .Saturation = (720 * val) / 360

                    End Select

                End With

                Exit Sub

            End If

            Select Case Mx

                Case r
                    hue = ((g - b) / chroma) Mod 6

                Case g
                    hue = ((b - r) / chroma) + 2

                Case b
                    hue = ((r - g) / chroma) + 4

            End Select

            hue *= 60
            If (hue < 0) Then hue = 360 + hue
            sat = chroma / val

            hsv.Value = val
            hsv.Hue = hue
            hsv.Saturation = sat

        End Sub

        Public Shared Function HSVToColor(hsv As HSVDATA) As System.Drawing.Color
            Dim c = HSVToMediaColor(hsv)
            Return Color.FromArgb(c.A, c.R, c.G, c.B)
        End Function

        Public Shared Function HSVToMediaColor(hsv As HSVDATA) As System.Windows.Media.Color

            '' http://en.wikipedia.org/wiki/HSL_and_HSV#Hue_and_chroma
            '' I adapted the equation from the one in Wikipedia.
            '' I wish I could offer a better explanation.  But this isn't wikipedia, and they'd do a better job.
            ''
            On Error Resume Next

            Dim a As Double,
            b As Double,
            c As Double

            Dim chroma As Double,
            Mx As Double,
            Mn As Double

            Dim j As System.Windows.Media.Color

            Dim n As Double

            If (hsv.Hue >= 360) Then hsv.Hue -= 360

            If (hsv.Hue = -1) Then

                With hsv
                    If (.Saturation) > (.Value) Then
                        .Saturation = 1
                        n = ((.Value * 360) / 510)
                    Else
                        n = 1 - ((.Saturation * 360) / 720)
                    End If
                End With

                Return System.Windows.Media.Color.FromScRgb(1, n, n, n)
            End If

            chroma = hsv.Value * hsv.Saturation

            Mn = Math.Abs(hsv.Value - chroma)
            Mx = hsv.Value

            n = hsv.Hue / 60

            a = Mx
            c = Mn

            b = chroma * (1 - Math.Abs((n Mod 2) - 1))

            b += c

            '' fit the color space in to byte space.

            ' Get the floored value of n
            n = Math.Floor(n)
            Select Case n

                Case 0, 6 '' 0, 360 - Red
                    j = System.Windows.Media.Color.FromScRgb(1, a, b, c)

                Case 1 '' 60 - Yellow
                    j = System.Windows.Media.Color.FromScRgb(1, b, a, c)

                Case 2 '' 120 - Green
                    j = System.Windows.Media.Color.FromScRgb(1, c, a, b)

                Case 3 '' 180 - Cyan
                    j = System.Windows.Media.Color.FromScRgb(1, c, b, a)

                Case 4 '' 240 - Blue
                    j = System.Windows.Media.Color.FromScRgb(1, b, c, a)

                Case 5 '' 300 - Magenta
                    j = System.Windows.Media.Color.FromScRgb(1, a, c, b)

            End Select

            Return j

        End Function

        Public Shared Function HSVToColorRaw(hsv As HSVDATA) As Integer

            '' http://en.wikipedia.org/wiki/HSL_and_HSV#Hue_and_chroma
            '' I adapted the equation from the one in Wikipedia.
            '' I wish I could offer a better explanation.  But this isn't wikipedia, and they'd do a better job.
            ''
            On Error Resume Next

            Dim a As Double,
            b As Double,
            c As Double

            Dim chroma As Double,
            Mx As Double,
            Mn As Double

            Dim j As Integer = &HFF000000

            Dim n As Double

            If (hsv.Hue >= 360) Then hsv.Hue -= 360

            If (hsv.Hue = -1) Then

                With hsv
                    If (.Saturation) > (.Value) Then
                        .Saturation = 1
                        n = ((.Value * 360) / 510)
                    Else
                        n = 1 - ((.Saturation * 360) / 720)
                    End If

                End With

                a = n * 255
                Return &HFF Or (a << 16) Or (a << 8) Or (a)

            End If

            chroma = hsv.Value * hsv.Saturation

            Mn = Math.Abs(hsv.Value - chroma)
            Mx = hsv.Value

            n = hsv.Hue / 60

            a = Mx
            c = Mn

            b = chroma * (1 - Math.Abs((n Mod 2) - 1))

            b += c

            '' fit the color space in to byte space.
            a = Math.Round(a * 255)
            b = Math.Round(b * 255)
            c = Math.Round(c * 255)

            ' Get the floored value of n
            n = Math.Floor(n)
            Select Case n

                Case 0, 6 '' 0, 360 - Red
                    j = j Or (a << 16) Or (b << 8) Or c

                Case 1 '' 60 - Yellow
                    j = j Or (b << 16) Or (a << 8) Or c

                Case 2 '' 120 - Green
                    j = j Or (c << 16) Or (a << 8) Or b

                Case 3 '' 180 - Cyan
                    j = j Or (c << 16) Or (b << 8) Or a

                Case 4 '' 240 - Blue
                    j = j Or (b << 16) Or (c << 8) Or a

                Case 5 '' 300 - Magenta
                    j = j Or (a << 16) Or (c << 8) Or b

            End Select

            HSVToColorRaw = j
        End Function

        Public Shared Sub ColorToCMY(ByVal Color As System.Drawing.Color, ByRef cmy As CMYDATA)
            ''
            Dim r As Integer,
            g As Integer,
            b As Integer

            Dim x As Integer

            GetRGB(Color, r, g, b)

            x = Max(r, g, b)

            r = Abs(1 - r)
            g = Abs(1 - g)
            b = Abs(1 - b)

            cmy.Magenta = r
            cmy.Yellow = g
            cmy.Cyan = b

        End Sub

        Public Shared Function CMYToColor(cmy As CMYDATA) As System.Drawing.Color
            ''
            Dim c As Integer,
            m As Integer,
            y As Integer

            Dim r As Integer,
            g As Integer,
            b As Integer

            c = cmy.Cyan
            m = cmy.Magenta
            y = cmy.Yellow

            r = Abs((1 - m))
            g = Abs((1 - y))
            b = Abs((1 - c))

            CMYToColor = System.Drawing.Color.FromArgb(255, r, g, b)

        End Function

        Public Shared Sub GetPercentages(ByVal Color As System.Drawing.Color, ByRef dpRed As Double, ByRef dpGreen As Double, ByRef dpBlue As Double)

            Dim vR As Double,
            vB As Double,
            vG As Double

            Dim vIn(0 To 2) As Byte
            Dim d As Double

            GetRGB(Color, vIn(0), vIn(1), vIn(2))

            vR = CDbl(vIn(0))
            vG = CDbl(vIn(1))
            vB = CDbl(vIn(2))

            If (vR > vG) Then d = vR Else d = vG
            If (vB > d) Then d = vB

            If (d = 0&) Then d = 255.0#

            dpRed = (vR / d)
            dpGreen = (vG / d)
            dpBlue = (vB / d)

        End Sub

        Public Shared Function AbsTone(ByVal Color As System.Drawing.Color) As System.Drawing.Color

            Dim pR As Double,
            pB As Double,
            pG As Double

            Dim sR As Double,
            sB As Double,
            sG As Double

            GetPercentages(Color, pR, pG, pB)

            sR = (pR) * 255
            sG = (pG) * 255
            sB = (pB) * 255

            AbsTone = System.Drawing.Color.FromArgb(Color.A, CInt(sR), CInt(sG), CInt(sB))

        End Function

        Public Overloads Shared Function SetTone(ByRef rgbData As RGBDATA, ByVal pPercent As Single, Optional ByVal Color As System.Drawing.Color = Nothing) As System.Drawing.Color

            Dim x As Single

            If (Color <> Nothing) Then ColorToRGB(Color, rgbData)
            x = Max(rgbData.Red, rgbData.Green, rgbData.Blue)

            If (x = 0) Then
                rgbData.Red = 0
                rgbData.Green = 0
                rgbData.Blue = 0

                Return System.Drawing.Color.Empty
            End If

            rgbData.Red = CByte((CSng(rgbData.Red) / x) * (255 * pPercent))
            rgbData.Green = CByte((CSng(rgbData.Green) / x) * (255 * pPercent))
            rgbData.Blue = CByte((CSng(rgbData.Blue) / x) * (255 * pPercent))

            Return RGBToColor(rgbData)
        End Function

        Public Overloads Shared Function SetTone(ByRef argbData As ARGBDATA, ByVal pPercent As Single, Optional ByVal Color As System.Drawing.Color = Nothing) As System.Drawing.Color

            Dim x As Single

            If (Color <> Nothing) Then ColorToARGB(Color, argbData)
            x = Max(argbData.Red, argbData.Green, argbData.Blue)

            If (x = 0) Then
                argbData.Red = 0
                argbData.Green = 0
                argbData.Blue = 0

                Return System.Drawing.Color.Empty
            End If

            argbData.Red = CByte((CSng(argbData.Red) / x) * (255 * pPercent))
            argbData.Green = CByte((CSng(argbData.Green) / x) * (255 * pPercent))
            argbData.Blue = CByte((CSng(argbData.Blue) / x) * (255 * pPercent))

            Return ARGBToColor(argbData)
        End Function

        Public Shared Function GrayTone(ByVal Color As System.Drawing.Color) As System.Drawing.Color

            Dim rgbData As ARGBDATA,
            a As Integer,
            b As Integer,
            c As Integer

            ColorToARGB(Color, rgbData)

            a = rgbData.Red
            b = rgbData.Green
            c = rgbData.Blue

            a = (a + b + c) / 3

            GrayTone = System.Drawing.Color.FromArgb(rgbData.Alpha, a, a, a)

        End Function

        Public Shared Function GetAverageColor(ByVal Color1 As System.Drawing.Color, ByVal Color2 As System.Drawing.Color) As System.Drawing.Color
            Dim Bits(0 To 2) As RGBDATA,
            df As Single,
            e As Integer

            Dim clr2 As Single = CSng(Color2.ToArgb())

            Dim al As Byte = CByte(Color1.A)
            Dim af As Byte = CByte(Color2.A)

            ColorToRGB(Color1, Bits(0))
            ColorToRGB(Color2, Bits(1))

            e = 0

            If (Math.Round(clr2) <> clr2) Then
                df = (CSng(Bits(0).Red) * clr2)
                e = CInt(Math.Round(df, 0))
                Bits(2).Red = CByte(e And &HFF)

                df = (CSng(Bits(0).Green) * clr2)
                e = CInt(Math.Round(df, 0))
                Bits(2).Green = CByte(e And &HFF)

                df = (CSng(Bits(0).Blue) * clr2)
                e = CInt(Math.Round(df, 0))
                Bits(2).Blue = CByte(e And &HFF)
            Else
                df = (CSng(Bits(0).Red) + CSng(Bits(1).Red)) / 2&
                e = CInt(Math.Round(df, 0))
                Bits(2).Red = CByte(e And &HFF)

                df = (CSng(Bits(0).Green) + CSng(Bits(1).Green)) / 2&
                e = CInt(Math.Round(df, 0))
                Bits(2).Green = CByte(e And &HFF)

                df = (CSng(Bits(0).Blue) + CSng(Bits(1).Blue)) / 2&
                e = CInt(Math.Round(df, 0))
                Bits(2).Blue = CByte(e And &HFF)
            End If

            '' Get the average alpha
            al = CByte((CInt(al) + CInt(af)) / 2)

            Return System.Drawing.Color.FromArgb(al, Bits(2).Red, Bits(2).Green, Bits(2).Blue)

        End Function

    End Class

End Namespace