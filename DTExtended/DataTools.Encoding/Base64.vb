Option Explicit On
Imports DataTools.ExtendedMath

Namespace Encoding

    Public Structure BASE64STRUCT
        Public Length As Integer
        Public Data() As Byte
    End Structure

    Public NotInheritable Class Base64

        Public Const BASE64TABLE = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
        Public Const BASE64PAD = 61
        Public Const BASE64PADRETURN = 254

        Public Shared B64CodeOut(0 To 63) As Byte
        Public Shared B64CodeReturn(0 To 255) As Byte
        Public Shared B64TableCreated As Boolean

        Private Sub New()

        End Sub

        Shared Sub New()
            Dim i As Integer, d As Integer

            For i = 0 To 255
                B64CodeReturn(i) = &H7F
            Next i

            For i = 0 To 63
                d = Asc(Mid(BASE64TABLE, i + 1, 1))
                B64CodeOut(i) = d And 255
                B64CodeReturn(d) = i And &H3F
            Next i

            B64CodeReturn(BASE64PAD) = BASE64PADRETURN

            B64TableCreated = True
        End Sub


        Public Shared Function Decode64(ByVal DataIn() As Byte, ByVal Length As Integer, ByRef DataOut() As Byte, Optional ByVal ShowProgress As Boolean = True, Optional ByVal ProgressString As String = "") As Integer
            On Error Resume Next

            Dim p As Integer

            Dim l As Integer,
            n As Integer,
            j As Integer,
            v As Integer

            Dim Quartet(0 To 3) As Byte

            If (ShowProgress = True) Then

            End If

            l = (Length / 4) * 3
            ReDim DataOut(0 To l)

            v = 0
            n = 0
            For l = 0 To (Length - 1) Step 0

                j = 0
                Do While (j < 4)

                    If (B64CodeReturn(DataIn(l)) <> &H7F) Then

                        Quartet(j) = B64CodeReturn(DataIn(l))
                        j = j + 1


                    End If

                    l = l + 1

                    If (l >= Length) Then

                        Do While j < 4
                            Quartet(j) = BASE64PADRETURN
                            j = j + 1
                        Loop

                        Exit Do
                    End If

                Loop

                If (Quartet(0) = BASE64PADRETURN) Or (Quartet(1) = BASE64PADRETURN) Then
                    l = -1&
                    Exit For
                End If

                DataOut(v) = LShift(Quartet(0), 2) Or RShift(Quartet(1), 4)
                v = v + 1

                If (Quartet(2) = BASE64PADRETURN) Then
                    l = -1&
                    Exit For
                End If

                DataOut(v) = LShift(Quartet(1), 4) Or RShift(Quartet(2), 2)
                v = v + 1

                If (Quartet(3) = BASE64PADRETURN) Then
                    l = -1&
                    Exit For
                End If

                DataOut(v) = LShift(Quartet(2), 6) Or Quartet(3)
                v = v + 1

                If ShowProgress = True Then
                    If n >= (Length / ProgressDiv) Then
                        n = 0
                        p = (l / (Length - 1)) * 100

                        ' Application.DoEvents()
                    Else
                        n = n + 1
                    End If
                End If

            Next l

            ReDim Preserve DataOut(0 To (v - 1))

            Decode64 = v

            If ShowProgress = True Then

            End If

        End Function

        Public Shared Function Encode64(ByVal DataIn() As Byte, ByVal Length As Integer, ByRef b64Out As BASE64STRUCT, Optional ByVal ShowProgress As Boolean = True, Optional ByVal ProgressString As String = "") As Boolean

            Dim p As Integer

            Dim l As Integer,
            x As Integer,
            v As Integer

            Dim lC As Integer,
            n As Integer

            Dim lProcess As Integer,
            lReturn As Integer,
            lActual As Integer

            If (ShowProgress = True) Then

            End If

            lActual = Length
            lProcess = Length
            If (lProcess Mod 3) Then
                lProcess = lProcess + (3 - (lProcess Mod 3))
            End If

            lReturn = (lProcess / 3) * 4

            ReDim Preserve DataIn(0 To (lProcess - 1))
            If Math.Round(lReturn / 76) <> (lReturn / 76) Then
                v = Math.Round(lReturn / 76) + 1
            Else
                v = lReturn / 76
            End If

            v = v * 2

            l = (lReturn - 1) + v

            ReDim b64Out.Data(0 To l)
            b64Out.Length = l + 1

            l = lProcess - 1
            v = 0

            If ShowProgress = True Then


            End If

            n = 0
            For x = 0 To l Step 3

                b64Out.Data(v) = B64CodeOut(RShift(DataIn(x), 2))
                b64Out.Data(v + 1) = B64CodeOut((LShift(DataIn(x), 4) And &H30) Or
                                 (RShift(DataIn(x + 1), 4)))
                b64Out.Data(v + 2) = B64CodeOut((LShift(DataIn(x + 1), 2) And &H3C) Or
                                  (RShift(DataIn(x + 2), 6)))

                b64Out.Data(v + 3) = B64CodeOut(DataIn(x + 2) And &H3F)

                v = v + 4
                lC = lC + 4
                If (lC >= 76) Then

                    lC = 0
                    If (l - x) > 3 Then
                        b64Out.Data(v) = 13
                        b64Out.Data(v + 1) = 10
                        v = v + 2
                    End If
                End If

                If ShowProgress = True Then
                    If n >= (l / ProgressDiv) Then
                        n = 0
                        p = (x / l) * 100

                        ' Application.DoEvents()
                    Else
                        n = n + 1
                    End If
                End If

            Next x

            Select Case (lProcess - lActual)

                Case 1
                    b64Out.Data(v - 1) = BASE64PAD

                Case 2
                    b64Out.Data(v - 1) = BASE64PAD
                    b64Out.Data(v - 2) = BASE64PAD

            End Select

            If lC <> 0 Then
                b64Out.Data(v) = 13
                b64Out.Data(v + 1) = 10
                ReDim Preserve b64Out.Data(0 To v + 1)

            Else
                ReDim Preserve b64Out.Data(0 To v - 1)

            End If

            If (ShowProgress = True) Then

            End If

            Return b64Out.Length

        End Function

    End Class

End Namespace