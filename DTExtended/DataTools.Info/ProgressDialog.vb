
Imports System.Windows.Forms

Namespace Info

    Public Module ProgressTools
        Public gAllProgress As New Collection

        Public Sub AddProgress(ByVal o As Object)
            Dim v

            For Each v In gAllProgress
                If v Is o Then
                    Return
                End If
            Next

            gAllProgress.Add(o)
        End Sub


        Public Sub RemoveProgress(ByVal o As Object)
            Dim i As Integer = 1
            For Each v As ProgressDialog In gAllProgress
                If v Is o Then
                    gAllProgress.Remove(i)
                    Return
                End If
                i += 1
            Next
        End Sub

        Public Sub CloseAllProgress()
            Dim v

            For Each v In gAllProgress
                v.close()
            Next
            gAllProgress.Clear()
        End Sub

    End Module

    Public Class ProgressDialog

        Public Sub New()

            ' This call is required by the Windows Form Designer.
            InitializeComponent()

            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = 100
            ProgressBar1.Value = 0
            ProgressBar1.Style = ProgressBarStyle.Blocks
            AddProgress(Me)

            ' Add any initialization after the InitializeComponent() call.

        End Sub


        Public Property Style As ProgressBarStyle
            Get
                Return ProgressBar1.Style
            End Get
            Set(value As ProgressBarStyle)
                ProgressBar1.Style = value
            End Set
        End Property

        Public Property Caption As String
            Get
                Return lblCaption.Text
            End Get
            Set(value As String)
                lblCaption.Text = value
            End Set
        End Property

        Public Property Value As Integer
            Get
                Return ProgressBar1.Value
            End Get
            Set(value As Integer)
                ProgressBar1.Value = value
            End Set
        End Property

        Public Property MaxValue As Integer
            Get
                Return ProgressBar1.Maximum
            End Get
            Set(value As Integer)
                ProgressBar1.Maximum = value
            End Set
        End Property

        Public Property MinValue As Integer
            Get
                Return ProgressBar1.Minimum
            End Get
            Set(value As Integer)
                ProgressBar1.Minimum = value
            End Set
        End Property

        Public Sub CloseOut()
            RemoveProgress(Me)

            Me.Close()
        End Sub
    End Class

End Namespace