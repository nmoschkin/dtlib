'' ************************************************* ''
'' DataTools Visual Basic Utility Library 
''
'' Module: Prints and Interprets Numeric Values
''         From Any Printed Base.
'' 
'' Copyright (C) 2011-2020 Nathan Moschkin
'' All Rights Reserved
''
'' Licensed Under the Microsoft Public License   
'' ************************************************* ''

Namespace Strings

    ''' <summary>
    ''' Multiple base number printing using alphanumeric characters: Up to base 62 (by default).
    ''' </summary>
    ''' <remarks></remarks>
    Public Module Strings

        Enum PadTypes
            None = 0
            [Byte] = 1
            [Short] = 2
            [Int] = 4
            [Long] = 8
            Auto = 10
        End Enum

        Private Function MakeBase(ByVal Number As Long, Optional workChars As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz") As String
            If (Number > workChars.Length) Or (Number < 2) Then Throw New ArgumentException("workChars", "Number of working characters does not meet or exceed the desired base.")
            Return Mid(workChars, 1, Number)
        End Function

        ''' <summary>
        ''' Returns a number from a string of the given base.
        ''' </summary>
        ''' <param name="ValueString">The numeric value string to parse.</param>
        ''' <param name="Base">The base to use in order to parse the string.</param>
        ''' <param name="workChars">Specifies an alternate set of glyphs to use for translation.</param>
        ''' <returns>A 64 bit unsigned number.</returns>
        ''' <remarks></remarks>
        Public Function MBValue(ByVal ValueString As String, Optional ByVal Base As Integer = 10, Optional workChars As String = Nothing) As ULong
            Dim i As ULong,
                b As ULong

            Dim j As ULong,
                mbStr As String,
                s As String

            Dim x As String

            b = Base
            mbStr = MakeBase(Base)
            If (mbStr = "") Then Return 0

            s = ValueString
            b = s.Length

            For j = 1 To b
                x = Mid(s, j, 1)

                i = (i * Base) + (InStr(1, mbStr, x) - 1)
            Next j

            Return i

        End Function

        ''' <summary>
        ''' Prints a number as a string of the given base.
        ''' </summary>
        ''' <param name="value">The value to print (must be an integer type of 8, 16, 32 or 64 bits; floating point values are not allowed).</param>
        ''' <param name="Base">Specifies the numeric base used to calculated the printed characters.</param>
        ''' <param name="PadType">Specifies the type of padding to use.</param>
        ''' <param name="workChars">Specifies an alternate set of glyphs to use for printing.</param>
        ''' <returns>A character string representing the input value as printed text in the desired base.</returns>
        ''' <remarks></remarks>
        Public Function MBString(ByVal value As Object, Optional ByVal Base As Integer = 10, Optional ByVal PadType As PadTypes = PadTypes.Auto, Optional workChars As String = Nothing) As String
            Dim varWork As ULong

            Dim i As ULong,
                b As ULong

            Dim j As ULong,
                mbStr As String,
                s As String

            Dim sLen As Integer = 0

            On Error Resume Next

            If PadType = PadTypes.Auto Then
                Select Case value.GetType
                    Case GetType(Long), GetType(ULong)
                        PadType = PadTypes.Long

                    Case GetType(Integer), GetType(UInteger)
                        PadType = PadTypes.Int

                    Case GetType(Short), GetType(UShort)
                        PadType = PadTypes.Short

                    Case GetType(Byte), GetType(SByte)
                        PadType = PadTypes.Byte

                    Case Else
                        Return ""

                End Select
            End If

            Select Case PadType
                Case PadTypes.Long
                    sLen = GetMaxPadLong(Base)

                Case PadTypes.Int
                    sLen = GetMaxPadInt(Base)

                Case PadTypes.Short
                    sLen = GetMaxPadShort(Base)

                Case PadTypes.Byte
                    sLen = GetMaxPadByte(Base)

            End Select

            varWork = Math.Abs(value)

            b = Base

            If workChars IsNot Nothing AndAlso workChars.Length = 0 Then
                workChars = Nothing
            End If

            If workChars Is Nothing Then
                mbStr = MakeBase(Base)
            Else
                mbStr = MakeBase(Base, workChars)
            End If

            If (mbStr Is Nothing) Then Throw New ArgumentException("workChars", "Cannot work with a null glyph set.")
            s = ""

            Do While varWork > 0
                If (varWork >= b) Then
                    j = (varWork Mod b)
                    j = j + 1
                Else
                    j = varWork + 1
                End If

                s = Mid(mbStr, j, 1) & s

                If (varWork < b) Then Exit Do

                varWork = (varWork - (j - 1)) / b
            Loop

            If (sLen) AndAlso (sLen - s.Length) > 0 Then
                s = New String(Mid(mbStr, 1, 1), sLen - s.Length) + s
            End If

            MBString = s

        End Function

        ''' <summary>
        ''' Calculate the maximum number of glyphs needed to represent a 64-bit number of the given base.
        ''' </summary>
        ''' <param name="Base"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetMaxPadLong(ByVal Base As Integer) As Integer
            Dim sVal As ULong

            sVal = &HFFFFFFFFFFFFFFFFUL
            GetMaxPadLong = Len(MBString(sVal, Base, PadTypes.None))
        End Function

        ''' <summary>
        ''' Calculate the maximum number of glyphs needed to represent a 32-bit number of the given base.
        ''' </summary>
        ''' <param name="Base"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetMaxPadInt(ByVal Base As Integer) As Integer
            Dim sVal As UInteger

            sVal = &HFFFFFFFFUI
            Return Len(MBString(sVal, Base, PadTypes.None))
        End Function

        ''' <summary>
        ''' Calculate the maximum number of glyphs needed to represent a 16-bit number of the given base.
        ''' </summary>
        ''' <param name="Base"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetMaxPadShort(ByVal Base As Integer) As Integer
            Dim sVal As UShort

            sVal = &HFFFFUS
            Return Len(MBString(sVal, Base, PadTypes.None))
        End Function

        ''' <summary>
        ''' Calculate the maximum number of glyphs needed to represent a 8-bit number of the given base.
        ''' </summary>
        ''' <param name="Base"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetMaxPadByte(ByVal Base As Integer) As Integer
            Dim sVal As Byte

            sVal = &HFF
            Return Len(MBString(sVal, Base, PadTypes.None))
        End Function

    End Module

End Namespace