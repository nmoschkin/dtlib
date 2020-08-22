Namespace Persistence

    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
    Partial Class GuidEditorDialog
        Inherits System.Windows.Forms.Form

        'Form overrides dispose to clean up the component list.
        <System.Diagnostics.DebuggerNonUserCode()>
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            Try
                If disposing AndAlso components IsNot Nothing Then
                    components.Dispose()
                End If
            Finally
                MyBase.Dispose(disposing)
            End Try
        End Sub

        'Required by the Windows Form Designer
        Private components As System.ComponentModel.IContainer

        'NOTE: The following procedure is required by the Windows Form Designer
        'It can be modified using the Windows Form Designer.  
        'Do not modify it using the code editor.
        <System.Diagnostics.DebuggerStepThrough()>
        Private Sub InitializeComponent()
            Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
            Me.OK_Button = New System.Windows.Forms.Button()
            Me.Cancel_Button = New System.Windows.Forms.Button()
            Me.btnGenerate = New System.Windows.Forms.Button()
            Me.tbGuid = New System.Windows.Forms.TextBox()
            Me.Label1 = New System.Windows.Forms.Label()
            Me.TableLayoutPanel1.SuspendLayout()
            Me.SuspendLayout()
            '
            'TableLayoutPanel1
            '
            Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.TableLayoutPanel1.ColumnCount = 2
            Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
            Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
            Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
            Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
            Me.TableLayoutPanel1.Location = New System.Drawing.Point(236, 66)
            Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
            Me.TableLayoutPanel1.RowCount = 1
            Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
            Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
            Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
            Me.TableLayoutPanel1.TabIndex = 0
            '
            'OK_Button
            '
            Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
            Me.OK_Button.Location = New System.Drawing.Point(3, 3)
            Me.OK_Button.Name = "OK_Button"
            Me.OK_Button.Size = New System.Drawing.Size(67, 23)
            Me.OK_Button.TabIndex = 4
            Me.OK_Button.Text = "OK"
            '
            'Cancel_Button
            '
            Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
            Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
            Me.Cancel_Button.Name = "Cancel_Button"
            Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
            Me.Cancel_Button.TabIndex = 5
            Me.Cancel_Button.Text = "Cancel"
            '
            'btnGenerate
            '
            Me.btnGenerate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.btnGenerate.Location = New System.Drawing.Point(312, 23)
            Me.btnGenerate.Name = "btnGenerate"
            Me.btnGenerate.Size = New System.Drawing.Size(67, 23)
            Me.btnGenerate.TabIndex = 3
            Me.btnGenerate.Text = "&Generate"
            '
            'tbGuid
            '
            Me.tbGuid.Location = New System.Drawing.Point(55, 25)
            Me.tbGuid.Name = "tbGuid"
            Me.tbGuid.Size = New System.Drawing.Size(251, 20)
            Me.tbGuid.TabIndex = 2
            '
            'Label1
            '
            Me.Label1.AutoSize = True
            Me.Label1.Location = New System.Drawing.Point(12, 28)
            Me.Label1.Name = "Label1"
            Me.Label1.Size = New System.Drawing.Size(37, 13)
            Me.Label1.TabIndex = 1
            Me.Label1.Text = "GUID:"
            '
            'GuidEditorDialog
            '
            Me.AcceptButton = Me.OK_Button
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.CancelButton = Me.Cancel_Button
            Me.ClientSize = New System.Drawing.Size(394, 107)
            Me.Controls.Add(Me.Label1)
            Me.Controls.Add(Me.tbGuid)
            Me.Controls.Add(Me.btnGenerate)
            Me.Controls.Add(Me.TableLayoutPanel1)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Name = "GuidEditorDialog"
            Me.ShowInTaskbar = False
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
            Me.Text = "Edit Guid"
            Me.TableLayoutPanel1.ResumeLayout(False)
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub
        Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
        Friend WithEvents OK_Button As System.Windows.Forms.Button
        Friend WithEvents Cancel_Button As System.Windows.Forms.Button
        Friend WithEvents btnGenerate As System.Windows.Forms.Button
        Friend WithEvents tbGuid As System.Windows.Forms.TextBox
        Friend WithEvents Label1 As System.Windows.Forms.Label

    End Class

End Namespace