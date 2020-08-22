Namespace Persistence

    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
    Partial Class VersionInfoDialog
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
            Me.btnAuto = New System.Windows.Forms.Button()
            Me.Cancel_Button = New System.Windows.Forms.Button()
            Me.OK_Button = New System.Windows.Forms.Button()
            Me.Label1 = New System.Windows.Forms.Label()
            Me._Maj = New System.Windows.Forms.NumericUpDown()
            Me._Min = New System.Windows.Forms.NumericUpDown()
            Me.Label2 = New System.Windows.Forms.Label()
            Me._Rev = New System.Windows.Forms.NumericUpDown()
            Me.Label3 = New System.Windows.Forms.Label()
            Me.TableLayoutPanel1.SuspendLayout()
            CType(Me._Maj, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me._Min, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me._Rev, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'TableLayoutPanel1
            '
            Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.TableLayoutPanel1.ColumnCount = 3
            Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33!))
            Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.34!))
            Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33!))
            Me.TableLayoutPanel1.Controls.Add(Me.btnAuto, 0, 0)
            Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 2, 0)
            Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 1, 0)
            Me.TableLayoutPanel1.Location = New System.Drawing.Point(105, 47)
            Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
            Me.TableLayoutPanel1.RowCount = 1
            Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
            Me.TableLayoutPanel1.Size = New System.Drawing.Size(219, 29)
            Me.TableLayoutPanel1.TabIndex = 0
            '
            'btnAuto
            '
            Me.btnAuto.Anchor = System.Windows.Forms.AnchorStyles.None
            Me.btnAuto.Location = New System.Drawing.Point(3, 3)
            Me.btnAuto.Name = "btnAuto"
            Me.btnAuto.Size = New System.Drawing.Size(66, 23)
            Me.btnAuto.TabIndex = 5
            Me.btnAuto.Text = "Auto"
            '
            'Cancel_Button
            '
            Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
            Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.Cancel_Button.Location = New System.Drawing.Point(148, 3)
            Me.Cancel_Button.Name = "Cancel_Button"
            Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
            Me.Cancel_Button.TabIndex = 3
            Me.Cancel_Button.Text = "Cancel"
            '
            'OK_Button
            '
            Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
            Me.OK_Button.Location = New System.Drawing.Point(75, 3)
            Me.OK_Button.Name = "OK_Button"
            Me.OK_Button.Size = New System.Drawing.Size(66, 23)
            Me.OK_Button.TabIndex = 4
            Me.OK_Button.Text = "OK"
            '
            'Label1
            '
            Me.Label1.AutoSize = True
            Me.Label1.Location = New System.Drawing.Point(10, 9)
            Me.Label1.Name = "Label1"
            Me.Label1.Size = New System.Drawing.Size(33, 13)
            Me.Label1.TabIndex = 1
            Me.Label1.Text = "Major"
            '
            '_Maj
            '
            Me._Maj.Location = New System.Drawing.Point(49, 7)
            Me._Maj.Name = "_Maj"
            Me._Maj.Size = New System.Drawing.Size(50, 20)
            Me._Maj.TabIndex = 2
            '
            '_Min
            '
            Me._Min.Location = New System.Drawing.Point(144, 7)
            Me._Min.Name = "_Min"
            Me._Min.Size = New System.Drawing.Size(57, 20)
            Me._Min.TabIndex = 4
            '
            'Label2
            '
            Me.Label2.AutoSize = True
            Me.Label2.Location = New System.Drawing.Point(105, 9)
            Me.Label2.Name = "Label2"
            Me.Label2.Size = New System.Drawing.Size(33, 13)
            Me.Label2.TabIndex = 3
            Me.Label2.Text = "Minor"
            '
            '_Rev
            '
            Me._Rev.Location = New System.Drawing.Point(261, 7)
            Me._Rev.Name = "_Rev"
            Me._Rev.Size = New System.Drawing.Size(59, 20)
            Me._Rev.TabIndex = 6
            '
            'Label3
            '
            Me.Label3.AutoSize = True
            Me.Label3.Location = New System.Drawing.Point(207, 9)
            Me.Label3.Name = "Label3"
            Me.Label3.Size = New System.Drawing.Size(48, 13)
            Me.Label3.TabIndex = 5
            Me.Label3.Text = "Revision"
            '
            'VersionInfoDialog
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(329, 88)
            Me.Controls.Add(Me._Rev)
            Me.Controls.Add(Me.Label3)
            Me.Controls.Add(Me._Min)
            Me.Controls.Add(Me.Label2)
            Me.Controls.Add(Me._Maj)
            Me.Controls.Add(Me.Label1)
            Me.Controls.Add(Me.TableLayoutPanel1)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Name = "VersionInfoDialog"
            Me.ShowInTaskbar = False
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
            Me.Text = "Version"
            Me.TableLayoutPanel1.ResumeLayout(False)
            CType(Me._Maj, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me._Min, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me._Rev, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub
        Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
        Friend WithEvents Label1 As System.Windows.Forms.Label
        Friend WithEvents _Maj As System.Windows.Forms.NumericUpDown
        Friend WithEvents _Min As System.Windows.Forms.NumericUpDown
        Friend WithEvents Label2 As System.Windows.Forms.Label
        Friend WithEvents _Rev As System.Windows.Forms.NumericUpDown
        Friend WithEvents Label3 As System.Windows.Forms.Label
        Friend WithEvents btnAuto As System.Windows.Forms.Button
        Friend WithEvents Cancel_Button As System.Windows.Forms.Button
        Friend WithEvents OK_Button As System.Windows.Forms.Button

    End Class

End Namespace