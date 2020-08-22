Imports System.Windows.Forms
Imports System.Drawing.Design
Imports System.ComponentModel

Namespace Persistence
    Public Class VersionInfoDialog

        Private WithEvents _VersionObject As VersionInfoObject = New VersionInfoObject

        Public Property Changed As Boolean

        Public Property VersionObject As VersionInfoObject
            Get
                Return _VersionObject
            End Get
            Set(value As VersionInfoObject)
                _VersionObject = value.Clone
                BindInit()
            End Set
        End Property

        Public Overloads Function ShowDialog(owner As IWin32Window, value As VersionInfoObject) As DialogResult

            VersionObject = value
            ShowDialog = ShowDialog(owner)

        End Function

        Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        End Sub

        Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.Close()
        End Sub

        Private Sub BindInit()
            _Maj.Value = 0

            _Maj.Minimum = 0
            _Maj.Maximum = 9999

            _Min.Value = 0

            _Min.Minimum = 0
            _Min.Maximum = 9999

            _Rev.Value = 0

            _Rev.Minimum = 0
            _Rev.Maximum = 99999

            Rebind()

        End Sub

        Private Sub Rebind()
            Dim b As New Binding("Value", _VersionObject, "Major")
            _Maj.DataBindings.Clear()
            _Maj.DataBindings.Add(b)

            b = New Binding("Value", _VersionObject, "Minor")
            _Min.DataBindings.Clear()
            _Min.DataBindings.Add(b)

            b = New Binding("Value", _VersionObject, "Revision")
            _Rev.DataBindings.Clear()
            _Rev.DataBindings.Add(b)

        End Sub
        Public Sub New(version As VersionInfoObject)
            ' This call is required by the designer.
            InitializeComponent()
            _VersionObject = version.Clone
            BindInit()
        End Sub

        Public Sub New()

            ' This call is required by the designer.
            InitializeComponent()
            BindInit()
            ' Add any initialization after the InitializeComponent() call.
        End Sub

        Private Sub btnAuto_Click(sender As Object, e As EventArgs) Handles btnAuto.Click
            _VersionObject.DateToMinRev()
            Rebind()
        End Sub

        Private Sub _VersionObject_PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Handles _VersionObject.PropertyChanged
            Changed = True
        End Sub

    End Class

    Public Class VersionObjectTypeEditor
        Inherits System.Drawing.Design.UITypeEditor

        Public Overrides Function GetEditStyle(context As System.ComponentModel.ITypeDescriptorContext) As System.Drawing.Design.UITypeEditorEditStyle
            'Return MyBase.GetEditStyle(context)
            Return UITypeEditorEditStyle.Modal

        End Function

        Public Overrides Function EditValue(context As System.ComponentModel.ITypeDescriptorContext, provider As System.IServiceProvider, value As Object) As Object

            Dim dlg As New VersionInfoDialog

            If dlg.ShowDialog(provider.GetService(GetType(System.Windows.Forms.IWin32Window)), value) = DialogResult.OK Then
                Return dlg.VersionObject
            Else
                Return value
            End If

        End Function

    End Class


End Namespace