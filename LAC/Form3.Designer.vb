<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Log_form
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Log_text = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'Log_text
        '
        Me.Log_text.Location = New System.Drawing.Point(25, 13)
        Me.Log_text.Multiline = True
        Me.Log_text.Name = "Log_text"
        Me.Log_text.Size = New System.Drawing.Size(843, 690)
        Me.Log_text.TabIndex = 0
        '
        'Log_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(889, 715)
        Me.Controls.Add(Me.Log_text)
        Me.Name = "Log_form"
        Me.Text = "Log"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Log_text As TextBox
End Class
