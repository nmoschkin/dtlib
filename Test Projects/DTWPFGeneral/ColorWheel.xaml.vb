Imports DataTools.Memory
Imports DataTools.ExtendedMath.ColorMath
Imports System.Drawing
Imports System.Drawing.Imaging

Public Class ColorWheel

    Public Function ProduceWheel(image As System.Drawing.Image) As System.Drawing.Image

        'Return image

        Dim wmaxbottom As Integer = -1

        Dim bm As New BitmapData
        Dim bm2 As New BitmapData
        Dim n As Bitmap = image

        If n.PixelFormat <> Imaging.PixelFormat.Format32bppArgb Then

            Dim tbmp = New Bitmap(n.Width, n.Height, Imaging.PixelFormat.Format32bppArgb)
            Dim g = Graphics.FromImage(tbmp)

            g.DrawImage(n, New Rectangle(0, 0, n.Width, n.Height))
            g.Dispose()
            n = tbmp
            'Return n
        End If
        'Return n

        n.LockBits(New Rectangle(0, 0, n.Width, n.Height), Imaging.ImageLockMode.ReadWrite, Imaging.PixelFormat.Format32bppArgb, bm)

        Dim b() As Integer
        Dim i As Integer,
            c As Integer

        ReDim b((bm.Width * bm.Height) - 1)

        '' take the unmanaged memory and make it something manageable and VB-like.
        MemCpy(b, bm.Scan0, bm.Stride * bm.Height)

        c = b.Length - 1

        Dim h As Integer = 0
        Dim hc As Integer = 0

        Dim tmargin As Integer


        For i = 0 To c

            If (b(i) And &HFFFFFF) <> &HFFFFFF Then
                Exit For
            End If

            hc += 1
            If hc = n.Width Then
                h += 1
                hc = 0
            End If

        Next

        tmargin = h

        h = n.Height
        hc = 0

        For i = c To 0 Step -1

            If (h < n.Height) Then
                If (b(i) And &HFFFFFF) <> &HFFFFFF Then
                    Exit For
                End If
            End If

            hc += 1
            If hc = n.Width Then
                h -= 1
                hc = 0
            End If

        Next

        If (h = 0) Then h = 1

        If wmaxbottom > h Then h = wmaxbottom

        If (n.Height - h >= tmargin) Then
            h += tmargin
        End If

        Dim bmp2 = New Bitmap(n.Width, h, Imaging.PixelFormat.Format32bppArgb)
        bmp2.LockBits(New Rectangle(0, 0, bmp2.Width, bmp2.Height), Imaging.ImageLockMode.ReadWrite, Imaging.PixelFormat.Format32bppArgb, bm2)

        MemCpy(bm2.Scan0, bm.Scan0, bm.Stride * h)

        bmp2.UnlockBits(bm2)
        n.UnlockBits(bm)

        Return bmp2

    End Function





End Class
