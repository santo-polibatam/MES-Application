<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DataGridUpdateDb
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
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Btn_update_Db = New System.Windows.Forms.Button()
        Me.txt_Search_db = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar()
        Me.btn_Add = New System.Windows.Forms.Button()
        Me.btn_refresh = New System.Windows.Forms.Button()
        Me.btn_Delete = New System.Windows.Forms.Button()
        Me.txtMaterial = New System.Windows.Forms.TextBox()
        Me.txtxDescription = New System.Windows.Forms.MaskedTextBox()
        Me.txtCategory = New System.Windows.Forms.MaskedTextBox()
        Me.txtQRCode = New System.Windows.Forms.MaskedTextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtReferences = New System.Windows.Forms.MaskedTextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.BackgroundColor = System.Drawing.Color.White
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(12, 12)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(686, 467)
        Me.DataGridView1.TabIndex = 0
        '
        'Btn_update_Db
        '
        Me.Btn_update_Db.Location = New System.Drawing.Point(12, 485)
        Me.Btn_update_Db.Name = "Btn_update_Db"
        Me.Btn_update_Db.Size = New System.Drawing.Size(75, 23)
        Me.Btn_update_Db.TabIndex = 1
        Me.Btn_update_Db.Text = "Update"
        Me.Btn_update_Db.UseVisualStyleBackColor = True
        '
        'txt_Search_db
        '
        Me.txt_Search_db.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_Search_db.Location = New System.Drawing.Point(520, 485)
        Me.txt_Search_db.Name = "txt_Search_db"
        Me.txt_Search_db.Size = New System.Drawing.Size(178, 22)
        Me.txt_Search_db.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(461, 490)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(47, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Search :"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripProgressBar1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 626)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(720, 22)
        Me.StatusStrip1.TabIndex = 4
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripProgressBar1
        '
        Me.ToolStripProgressBar1.Name = "ToolStripProgressBar1"
        Me.ToolStripProgressBar1.Size = New System.Drawing.Size(100, 16)
        '
        'btn_Add
        '
        Me.btn_Add.Location = New System.Drawing.Point(544, 37)
        Me.btn_Add.Name = "btn_Add"
        Me.btn_Add.Size = New System.Drawing.Size(95, 23)
        Me.btn_Add.TabIndex = 5
        Me.btn_Add.Text = "Add New Data"
        Me.btn_Add.UseVisualStyleBackColor = True
        '
        'btn_refresh
        '
        Me.btn_refresh.Location = New System.Drawing.Point(231, 485)
        Me.btn_refresh.Name = "btn_refresh"
        Me.btn_refresh.Size = New System.Drawing.Size(113, 23)
        Me.btn_refresh.TabIndex = 6
        Me.btn_refresh.Text = "Refresh Table"
        Me.btn_refresh.UseVisualStyleBackColor = True
        '
        'btn_Delete
        '
        Me.btn_Delete.Location = New System.Drawing.Point(99, 485)
        Me.btn_Delete.Name = "btn_Delete"
        Me.btn_Delete.Size = New System.Drawing.Size(120, 23)
        Me.btn_Delete.TabIndex = 7
        Me.btn_Delete.Text = "Delete Selected Row"
        Me.btn_Delete.UseVisualStyleBackColor = True
        '
        'txtMaterial
        '
        Me.txtMaterial.Location = New System.Drawing.Point(14, 40)
        Me.txtMaterial.Name = "txtMaterial"
        Me.txtMaterial.Size = New System.Drawing.Size(100, 20)
        Me.txtMaterial.TabIndex = 9
        '
        'txtxDescription
        '
        Me.txtxDescription.Location = New System.Drawing.Point(120, 40)
        Me.txtxDescription.Name = "txtxDescription"
        Me.txtxDescription.Size = New System.Drawing.Size(100, 20)
        Me.txtxDescription.TabIndex = 10
        '
        'txtCategory
        '
        Me.txtCategory.Location = New System.Drawing.Point(226, 40)
        Me.txtCategory.Name = "txtCategory"
        Me.txtCategory.Size = New System.Drawing.Size(100, 20)
        Me.txtCategory.TabIndex = 11
        '
        'txtQRCode
        '
        Me.txtQRCode.Location = New System.Drawing.Point(332, 40)
        Me.txtQRCode.Name = "txtQRCode"
        Me.txtQRCode.Size = New System.Drawing.Size(100, 20)
        Me.txtQRCode.TabIndex = 12
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(14, 23)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(44, 13)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "Material"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(120, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 13)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "Description"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(226, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(49, 13)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "Category"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(332, 24)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(51, 13)
        Me.Label5.TabIndex = 16
        Me.Label5.Text = "QR Code"
        '
        'txtReferences
        '
        Me.txtReferences.Location = New System.Drawing.Point(438, 40)
        Me.txtReferences.Name = "txtReferences"
        Me.txtReferences.Size = New System.Drawing.Size(100, 20)
        Me.txtReferences.TabIndex = 17
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(438, 24)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(62, 13)
        Me.Label6.TabIndex = 18
        Me.Label6.Text = "References"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(356, 485)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(99, 23)
        Me.Button1.TabIndex = 19
        Me.Button1.Text = "Delete Duplicate"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtxDescription)
        Me.GroupBox1.Controls.Add(Me.btn_Add)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.txtMaterial)
        Me.GroupBox1.Controls.Add(Me.txtReferences)
        Me.GroupBox1.Controls.Add(Me.txtCategory)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.txtQRCode)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 548)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(686, 67)
        Me.GroupBox1.TabIndex = 20
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Adding New Data"
        '
        'DataGridUpdateDb
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(720, 648)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.btn_Delete)
        Me.Controls.Add(Me.btn_refresh)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txt_Search_db)
        Me.Controls.Add(Me.Btn_update_Db)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "DataGridUpdateDb"
        Me.Text = "Update New data scan Components"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Btn_update_Db As Button
    Friend WithEvents txt_Search_db As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents ToolStripProgressBar1 As ToolStripProgressBar
    Friend WithEvents btn_Add As Button
    Friend WithEvents btn_refresh As Button
    Friend WithEvents btn_Delete As Button
    Friend WithEvents txtMaterial As TextBox
    Friend WithEvents txtxDescription As MaskedTextBox
    Friend WithEvents txtCategory As MaskedTextBox
    Friend WithEvents txtQRCode As MaskedTextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents txtReferences As MaskedTextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents GroupBox1 As GroupBox
End Class
