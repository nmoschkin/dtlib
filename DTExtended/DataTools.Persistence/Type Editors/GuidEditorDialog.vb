Imports System.Windows.Forms
Imports System.Windows.Forms.Design
Imports System.Drawing.Design
Imports System.Reflection
Imports System.Runtime.Serialization

Namespace Persistence

    Public Class GuidEditorDialog

        Private m_Guid As Guid

        Sub New()

            ' This call is required by the designer.
            InitializeComponent()
            tbGuid.Text = m_Guid.ToString

            ' Add any initialization after the InitializeComponent() call.

        End Sub

        Public Property Guid() As Guid
            Get
                Return m_Guid
            End Get
            Set(ByVal value As Guid)
                m_Guid = value
                tbGuid.Text = m_Guid.ToString
            End Set
        End Property

        Public Overloads Function ShowDialog(guid As Guid) As System.Windows.Forms.DialogResult

            Me.Guid = guid
            Return MyBase.ShowDialog()

        End Function

        Public Overloads Function ShowDialog(owner As System.Windows.Forms.IWin32Window, guid As Guid) As System.Windows.Forms.DialogResult

            Me.Guid = guid
            Return MyBase.ShowDialog(owner)

        End Function

        Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
            If ValidateGuid() = False Then Return

            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        End Sub

        Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.Close()
        End Sub

        Private Sub Generate(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenerate.Click
            NewGuid()
        End Sub

        Protected Sub NewGuid()
            Guid = Guid.NewGuid
        End Sub

        Protected Function ValidateGuid() As Boolean

            Dim g As Guid

            Try
                g = Guid.Parse(tbGuid.Text)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical)
                Return False
            End Try

            m_Guid = g
            Return True

        End Function

    End Class

    Public Class GuidTypeEditor
        Inherits System.Drawing.Design.UITypeEditor

        Public Overrides Function GetEditStyle(context As System.ComponentModel.ITypeDescriptorContext) As System.Drawing.Design.UITypeEditorEditStyle
            'Return MyBase.GetEditStyle(context)
            Return UITypeEditorEditStyle.Modal

        End Function

        Public Overrides Function EditValue(context As System.ComponentModel.ITypeDescriptorContext, provider As System.IServiceProvider, value As Object) As Object

            Dim dlg As New GuidEditorDialog

            If dlg.ShowDialog(provider.GetService(GetType(System.Windows.Forms.IWin32Window)), value) = DialogResult.OK Then
                Return dlg.Guid
            Else
                Return value
            End If

        End Function

    End Class

End Namespace