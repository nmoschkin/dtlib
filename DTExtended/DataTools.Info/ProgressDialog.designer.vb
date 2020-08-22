Namespace Info

    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
    Partial Class ProgressDialog
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
            Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
            Me.lblCaption = New System.Windows.Forms.Label
            Me.SuspendLayout()
            '
            'ProgressBar1
            '
            Me.ProgressBar1.Location = New System.Drawing.Point(12, 37)
            Me.ProgressBar1.Name = "ProgressBar1"
            Me.ProgressBar1.Size = New System.Drawing.Size(400, 23)
            Me.ProgressBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous
            Me.ProgressBar1.TabIndex = 0
            Me.ProgressBar1.UseWaitCursor = True
            '
            'lblCaption
            '
            Me.lblCaption.AutoSize = True
            Me.lblCaption.Location = New System.Drawing.Point(12, 9)
            Me.lblCaption.Name = "lblCaption"
            Me.lblCaption.Size = New System.Drawing.Size(48, 13)
            Me.lblCaption.TabIndex = 1
            Me.lblCaption.Text = "Progress"
            '
            'ProgressDialog
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(424, 72)
            Me.Controls.Add(Me.lblCaption)
            Me.Controls.Add(Me.ProgressBar1)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Name = "ProgressDialog"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.Text = "Progress"
            Me.UseWaitCursor = True
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub
        Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
        Friend WithEvents lblCaption As System.Windows.Forms.Label
    End Class

End Namespace