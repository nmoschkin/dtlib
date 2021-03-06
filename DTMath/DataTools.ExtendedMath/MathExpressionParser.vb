﻿Option Explicit On
Option Strict Off
Option Compare Binary

Imports System.Math

#Region "Math and Conversion"

Public Class MathExpressionParser

        Private m_Expression As String = ""

        Private m_Col As New Collection

        Private m_Error As String = ""
        Private m_Value As Double = 0.0#

        Public ReadOnly Property Value As Double
            Get
                Return m_Value
            End Get
        End Property

        Public Property ErrorText As String
            Get
                Return m_Error
            End Get
            Friend Set(value As String)
                m_Error = value
            End Set
        End Property

        Public Property Expression As String
            Get
                Return m_Expression
            End Get
            Set(value As String)
                m_Expression = value
            End Set
        End Property

        Public ReadOnly Property Operations As Collection
            Get
                Return m_Col
            End Get
        End Property

        Public Function ParseOperations(Optional value As String = Nothing) As Boolean

            Dim st() As String = Nothing,
            i As Integer,
            c As Integer

            Dim n As Integer = 0,
            b As Integer = 0

            Dim sLeft As String = "",
            sFar As String = ""

            Dim objOp As MathExpressionParser = Nothing

            If value = Nothing Then
                value = m_Expression
            End If

            value = OneSpace(SpaceOperators(Trim(value)).ToLower)
            m_Expression = value

            c = 0
            i = value.IndexOf("(")

            If i = -1 Then
                i = value.IndexOf(")")
                If i <> -1 Then
                    m_Error = "Unexpected closing parenthesis at column " & i
                    Return False
                End If

                m_Value = ParseExpression(value, m_Error)
                If m_Value = Nothing Then Return False

                Return True
            End If

            Do Until c >= value.Length
                i = value.IndexOf("(", c)
                If i = -1 Then
                    ReDim Preserve st(n)

                    st(n) = "@" & value.Substring(c)
                    n += 1

                    Exit Do
                End If

                ReDim Preserve st(n)

                If (i - c) > 0 Then
                    st(n) = "@" & value.Substring(c, (i - c))
                    n += 1
                End If

                sFar = Bracket(value, i, c)
                ReDim Preserve st(n)

                st(n) = sFar
                n += 1
                i = value.IndexOf("(", c)

            Loop

            n -= 1
            sFar = ""
            m_Col.Clear()

            For i = 0 To n
                ErrorText = ""

                If (st(i).Substring(0, 1) = "@") Then
                    sFar &= st(i).Substring(1)
                    objOp = New MathExpressionParser
                    objOp.ParseExpression(sFar)

                    m_Col.Add(objOp)
                Else
                    objOp = New MathExpressionParser
                    objOp.ParseOperations(st(i))

                    If (objOp.ErrorText <> "") Then
                        ErrorText = objOp.ErrorText
                        Return False
                    End If

                    m_Col.Add(objOp)
                    sFar &= " " & objOp.Value
                End If
            Next

            sFar = OneSpace(SpaceOperators(Trim(sFar)).ToLower)
            m_Value = ParseExpression(sFar, ErrorText)

            If ErrorText <> "" Then Return False
            Return True

        End Function

        Public Function ParseExpression(value As String, Optional ByRef ErrorText As String = Nothing) As Double
            Dim s() As String
            Dim dOut As Double = 0.0#,
            i As Integer = 0,
            j As Integer = 0,
            c As Integer

            Dim cc As Integer = 0,
            n As Integer = 0,
            m As Integer = 0

            Dim ob As Object

            Dim a As Long,
            b As Long

            Dim d As Double,
            e As Double

            Dim ops() As Object = Nothing,
            ops2() As Object = Nothing

            ' functions come first!
            Dim orderOps As String() = {"log", "log10", "sin", "cos", "tan", "asin", "acos", "atan", "^", "exp", "abs", "sqrt", "*", "/", "\", "mod", "%", "+", "-"}

            s = Words(Trim(OneSpace(SpaceOperators(value))))

            c = s.Length - 1

            For i = 0 To c

                ob = FVal(s(i))

                If ob = Nothing Then

                    If i = 0 Then

                        Select Case s(i)
                            Case "abs", "sqrt", "log", "log10", "sin", "cos", "tan", "asin", "acos", "atan"
                                Exit Select

                            Case Else
                                ErrorText = "Unexpected identifier '" & s(i) & "' at column " & cc + 1
                                Return Nothing

                        End Select
                    End If

                    Select Case s(i)

                        Case "^", "exp", "abs", "sqrt", "*", "/", "\", "mod", "%", "+", "-", "log", "log10", "sin", "cos", "tan", "asin", "acos", "atan"

                            ReDim Preserve ops(n)
                            ops(n) = s(i)
                            n += 1

                        Case Else
                            ErrorText = "Unexpected identifier '" & s(i) & "' at column " & cc + 1
                            Return Nothing

                    End Select
                Else
                    If s(i).Substring(0, 1) = "+" Then
                        ReDim Preserve ops(n)
                        ops(n) = "un"
                        n += 1
                    End If

                    ReDim Preserve ops(n)
                    ops(n) = ob
                    n += 1
                End If

                '' just to keep track of the character position in case there's an error
                cc += s(i).Length
            Next

            '' Now, perform the operations!

            c = orderOps.Length - 1
            n -= 1

            Try
                m = 0
                For j = 0 To n
                    If (j + 3 <= n) Then
                        If (IsNumeric(ops(j)) And IsNumeric(ops(j + 1))) Then

                            If ops(j + 2) = "/" And IsNumeric(ops(j + 3)) Then

                                ReDim Preserve ops2(m)
                                ops2(m) = ops(j) + (ops(j + 1) / ops(j + 3))
                                m += 1
                                j += 3

                            End If

                            Continue For
                        End If
                    End If

                    ReDim Preserve ops2(m)
                    ops2(m) = ops(j)
                    m += 1

                Next

                ops = ops2
                ops2 = Nothing
                n = m - 1

                For i = 0 To c

                    For j = 0 To n

                        If (Not TypeOf ops(j) Is String) Then
                            ReDim Preserve ops2(m)
                            ops2(m) = ops(j)
                            m += 1
                            Continue For
                        End If

                        If (ops(j) <> orderOps(i)) Then
                            ReDim Preserve ops2(m)
                            ops2(m) = ops(j)
                            m += 1
                            Continue For
                        End If

                        Select Case orderOps(i)

                            Case "^", "exp"

                                If j = 0 Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                If j = n Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If
                                m -= 1

                                d = ops(j - 1)
                                e = ops(j + 1)

                                If d < 0.0# Then
                                    d *= -1
                                    d = d ^ e
                                    d *= -1
                                Else
                                    If j - 1 > 0 Then
                                        If ops(j - 2) = "un" Then
                                            d = d ^ e
                                            d = Abs(d)
                                            m -= 1
                                        Else
                                            d = d ^ e
                                        End If
                                    Else
                                        d = d ^ e
                                    End If
                                End If

                                ReDim Preserve ops2(m)
                                ops2(m) = d
                                ops(j + 1) = ops2(m)
                                m += 1
                                j += 1
                                Continue For

                            Case "abs"

                                If j = n Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                d = Abs(ops(j + 1))

                                ReDim Preserve ops2(m)
                                ops2(m) = d
                                m += 1
                                j += 1

                                Continue For

                            Case "sqrt"
                                If j = n Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                d = Math.Sqrt(ops(j + 1))

                                ReDim Preserve ops2(m)
                                ops2(m) = d
                                m += 1
                                j += 1

                                Continue For

                            Case "*"
                                If j = n Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                If j = 0 Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                m -= 1

                                d = ops(j - 1)
                                e = ops(j + 1)

                                ReDim Preserve ops2(m)
                                ops2(m) = d * e
                                ops(j + 1) = ops2(m)
                                m += 1
                                j += 1

                                Continue For

                            Case "/"
                                If j = n Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                If j = 0 Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                m -= 1

                                d = ops(j - 1)
                                e = ops(j + 1)

                                ReDim Preserve ops2(m)
                                ops2(m) = d / e
                                ops(j + 1) = ops2(m)
                                m += 1
                                j += 1

                                Continue For

                            Case "\"
                                If j = n Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                If j = 0 Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                m -= 1

                                a = ops(j - 1)
                                b = ops(j + 1)

                                ReDim Preserve ops2(m)
                                ops2(m) = a \ b
                                ops(j + 1) = ops2(m)
                                m += 1
                                j += 1

                                Continue For

                            Case "mod", "%"
                                If j = n Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                If j = 0 Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                m -= 1

                                d = ops(j - 1)
                                e = ops(j + 1)

                                ReDim Preserve ops2(m)
                                ops2(m) = d Mod e
                                ops(j + 1) = ops2(m)
                                m += 1
                                j += 1

                                Continue For

                            Case "+"
                                If j = n Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                If j = 0 Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                m -= 1

                                d = ops(j - 1)
                                e = ops(j + 1)

                                ReDim Preserve ops2(m)
                                ops2(m) = d + e
                                ops(j + 1) = ops2(m)
                                m += 1
                                j += 1

                                Continue For

                            Case "-"
                                If j = n Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                If j = 0 Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                m -= 1

                                d = ops(j - 1)
                                e = ops(j + 1)

                                ReDim Preserve ops2(m)
                                ops2(m) = d - e
                                ops(j + 1) = ops2(m)
                                m += 1
                                j += 1

                                Continue For

                            Case "log"

                                If j = n Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                d = ops(j + 1)
                                d = Math.Log(d)
                                ReDim Preserve ops2(m)
                                ops2(m) = d
                                m += 1
                                j += 1

                                Continue For

                            Case "log10"
                                If j = n Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                d = ops(j + 1)
                                d = Math.Log10(d)
                                ReDim Preserve ops2(m)
                                ops2(m) = d
                                m += 1
                                j += 1

                                Continue For

                            Case "sin"

                                If j = n Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                d = ops(j + 1)
                                d = Math.Sin(d)
                                ReDim Preserve ops2(m)
                                ops2(m) = d
                                m += 1
                                j += 1

                                Continue For

                            Case "cos"

                                If j = n Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                If j = n Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                d = ops(j + 1)
                                d = Math.Cos(d)
                                ReDim Preserve ops2(m)
                                ops2(m) = d
                                m += 1
                                j += 1

                                Continue For

                            Case "tan"

                                If j = n Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If
                                If j = n Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                d = ops(j + 1)
                                d = Math.Tan(d)
                                ReDim Preserve ops2(m)
                                ops2(m) = d
                                m += 1
                                j += 1

                                Continue For

                            Case "asin"

                                If j = n Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                d = ops(j + 1)
                                d = Math.Asin(d)
                                ReDim Preserve ops2(m)
                                ops2(m) = d
                                m += 1
                                j += 1

                                Continue For

                            Case "acos"

                                If j = n Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                If j = n Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                d = ops(j + 1)
                                d = Math.Acos(d)
                                ReDim Preserve ops2(m)
                                ops2(m) = d
                                m += 1
                                j += 1

                                Continue For

                            Case "atan"

                                If j = n Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If
                                If j = n Then
                                    ErrorText = "Unexpected identifier '" & s(i)
                                    Return Nothing
                                End If

                                d = ops(j + 1)
                                d = Math.Atan(d)
                                ReDim Preserve ops2(m)
                                ops2(m) = d
                                m += 1
                                j += 1

                                Continue For

                        End Select

                    Next

                    ops = ops2
                    n = ops2.Length - 1
                    m = 0
                    ops2 = Nothing

                Next

                '' numbers that occur together add

                n = ops.Length - 1
                e = 0.0#
                For j = 0 To n
                    If IsNumeric(ops(j)) Then _
                    e += ops(j)
                Next

            Catch ex As Exception
                ErrorText = ex.Message
                Return Nothing
            End Try

            m_Error = ""
            m_Value = e

            Return e

        End Function
    End Class

#End Region

