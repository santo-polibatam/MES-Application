<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_preview
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form_preview))
        Me.pictureBoxPreview = New System.Windows.Forms.PictureBox()
        CType(Me.pictureBoxPreview, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pictureBoxPreview
        '
        Me.pictureBoxPreview.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pictureBoxPreview.Location = New System.Drawing.Point(0, 0)
        Me.pictureBoxPreview.Name = "pictureBoxPreview"
        Me.pictureBoxPreview.Size = New System.Drawing.Size(809, 519)
        Me.pictureBoxPreview.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.pictureBoxPreview.TabIndex = 0
        Me.pictureBoxPreview.TabStop = False
        '
        'Form_preview
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(809, 519)
        Me.Controls.Add(Me.pictureBoxPreview)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form_preview"
        Me.Text = "Preview"
        CType(Me.pictureBoxPreview, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents pictureBoxPreview As PictureBox
End Class
