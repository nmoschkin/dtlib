Imports System.IO

Namespace Encoding

    Public Class Base64Engine

        Public Enum DataTypeConstants
            dtBinary = &H0
            dtBase64 = &H1
        End Enum

        Private B64 As New BASE64STRUCT
        Private m_ShowProgress As Boolean = True

        Private m_DataType As DataTypeConstants

        Public ReadOnly Property DataType() As DataTypeConstants
            Get
                Return m_DataType
            End Get
        End Property

        Public ReadOnly Property Length() As Integer
            Get
                Return B64.Length
            End Get
        End Property

        Public Property ShowProgress() As Boolean
            Get
                Return m_ShowProgress
            End Get
            Set(ByVal value As Boolean)
                m_ShowProgress = value
            End Set
        End Property

        Public ReadOnly Property Data() As Byte()
            Get
                Data = B64.Data
            End Get
        End Property

        Public Property DataString(Optional ByVal Escape As Boolean = True) As String
            Get

                Return System.Text.Encoding.ASCII.GetString(B64.Data)

            End Get
            Set(ByVal v As String)

                B64.Data = System.Text.Encoding.ASCII.GetBytes(v)

            End Set
        End Property

        Public Sub Clear()
            Erase B64.Data
            B64.Length = 0

        End Sub

        Public Function GetDataBytes() As Byte()
            Return B64.Data

        End Function

        Public Function DecodeImage(ByVal s64 As String) As System.Drawing.Image

            Dim i As System.Drawing.Image,
            st As System.IO.MemoryStream

            'Me.DataString = s64

            'Try

            '    Me.Decode(Me.GetDataBytes(), Me.GetDataBytes().Length, "Decoding Byte Stream")

            'Catch ex As Exception
            '    MsgBox("Unknown error decoding file: " + ex.Message)
            '    Return Nothing
            'End Try


            'st = New System.IO.MemoryStream(Me.GetDataBytes)
            'i = Image.FromStream(st)

            'Return i

            DataString = s64

            Try

                Decode(GetDataBytes(), GetDataBytes().Length, "Decoding Byte Stream")

            Catch ex As Exception
                MsgBox("Unknown error decoding file: " + ex.Message)
                Return Nothing
            End Try


            st = New System.IO.MemoryStream(GetDataBytes)
            i = System.Drawing.Image.FromStream(st)

            Return i

        End Function


        ''' <summary>
        ''' Encodes a byte array into a Base64 byte array.
        ''' </summary>
        ''' <param name="input">The buffer to encode.</param>
        ''' <returns>The encoded buffer.</returns>
        ''' <remarks></remarks>
        Public Shared Function Base64Encode(input() As Byte) As Byte()
            Dim b As New BASE64STRUCT

            Base64.Encode64(input, input.Length, b, False)
            Return b.Data
        End Function


        ''' <summary>
        ''' Encodes a byte array into a Base64 string.
        ''' </summary>
        ''' <param name="input">The buffer to encode.</param>
        ''' <param name="encoding">The encoded string.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Base64Encode(input() As Byte, encoding As System.Text.Encoding) As String

            If encoding Is Nothing Then encoding = System.Text.Encoding.UTF8
            Return encoding.GetString(Base64Encode(input))

        End Function

        ''' <summary>
        ''' Decodes a Base64-encoded byte array.
        ''' </summary>
        ''' <param name="input">The input buffer.</param>
        ''' <returns>Decoded contents or nothing on error.</returns>
        ''' <remarks></remarks>
        Public Shared Function Base64Decode(input() As Byte) As Byte()

            Dim b() As Byte = Nothing

            If Base64.Decode64(input, input.Length, b, False) Then
                Return b
            End If

            Return Nothing

        End Function

        ''' <summary>
        ''' Decodes a Base64 input string into data.
        ''' </summary>
        ''' <param name="input">The string of Base64 characters to decode.</param>
        ''' <param name="encoding">The System.Text.Encoding option used to convert the string into an array of bytes.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Base64Decode(input As String, encoding As System.Text.Encoding) As Byte()

            Dim ib() As Byte = encoding.GetBytes(input)
            Dim b() As Byte = Nothing

            Base64.Decode64(ib, b.Length, b, False)

            Return b
        End Function


        Public Function Encode(ByVal Bytes() As Byte, ByVal Length As Integer, Optional ByVal ProgressString As String = "") As Integer
            Dim i As Integer
            Dim b() As Byte

            b = Bytes

            i = Base64.Encode64(b, Length, B64, m_ShowProgress, ProgressString)
            m_DataType = DataTypeConstants.dtBase64

            Return i
        End Function

        Public Function Decode(ByVal Bytes() As Byte, ByVal Length As Integer, Optional ByVal ProgressString As String = "") As Integer
            On Error Resume Next
            Dim b() As Byte

            b = Bytes

            Erase B64.Data

            ReDim B64.Data(0 To (Length - 1))
            B64.Length = Base64.Decode64(Bytes, Length, B64.Data, m_ShowProgress, ProgressString)
            m_DataType = DataTypeConstants.dtBinary

        End Function

        Public Overloads Function EncodeFile(ByVal fileName As String) As Integer

            Dim b() As Byte
            b = My.Computer.FileSystem.ReadAllBytes(fileName)

            Return Encode(b, b.Length)

        End Function

        Public Overloads Function EncodeFile(ByVal fileName As String, ByVal outputFile As String) As Integer

            Dim b() As Byte,
            l As Integer

            b = My.Computer.FileSystem.ReadAllBytes(fileName)

            l = Encode(b, b.Length)

            My.Computer.FileSystem.WriteAllBytes(outputFile, B64.Data, False)
            Return l

        End Function

        Public Overloads Function DecodeFile(ByVal fileName As String) As Integer
            Dim b() As Byte
            b = My.Computer.FileSystem.ReadAllBytes(fileName)

            Return Decode(b, b.Length)

        End Function

        Public Overloads Function DecodeFile(ByVal fileName As String, ByVal outputFile As String) As Integer
            Dim b() As Byte,
            l As Integer

            b = My.Computer.FileSystem.ReadAllBytes(fileName)

            l = Decode(b, b.Length)

            My.Computer.FileSystem.WriteAllBytes(outputFile, B64.Data, False)
            Return l

        End Function

    End Class

End Namespace