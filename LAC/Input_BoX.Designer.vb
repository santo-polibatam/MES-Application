<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Input_BoX
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Input_BoX))
        Me.Combo_PlanCode = New System.Windows.Forms.ComboBox()
        Me.Txt_Barcode = New System.Windows.Forms.TextBox()
        Me.Label165 = New System.Windows.Forms.Label()
        Me.Label164 = New System.Windows.Forms.Label()
        Me.Label163 = New System.Windows.Forms.Label()
        Me.txt_PlanCode_DateCode = New System.Windows.Forms.TextBox()
        Me.txt_QRCode = New System.Windows.Forms.TextBox()
        Me.btn_Cek = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Combo_PlanCode
        '
        Me.Combo_PlanCode.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo_PlanCode.FormattingEnabled = True
        Me.Combo_PlanCode.Items.AddRange(New Object() {"AF", "HH", "10"})
        Me.Combo_PlanCode.Location = New System.Drawing.Point(204, 118)
        Me.Combo_PlanCode.Name = "Combo_PlanCode"
        Me.Combo_PlanCode.Size = New System.Drawing.Size(48, 28)
        Me.Combo_PlanCode.TabIndex = 18
        Me.Combo_PlanCode.Visible = False
        '
        'Txt_Barcode
        '
        Me.Txt_Barcode.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Txt_Barcode.Location = New System.Drawing.Point(258, 118)
        Me.Txt_Barcode.Name = "Txt_Barcode"
        Me.Txt_Barcode.Size = New System.Drawing.Size(324, 26)
        Me.Txt_Barcode.TabIndex = 17
        Me.Txt_Barcode.Visible = False
        '
        'Label165
        '
        Me.Label165.AutoSize = True
        Me.Label165.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label165.Location = New System.Drawing.Point(23, 126)
        Me.Label165.Name = "Label165"
        Me.Label165.Size = New System.Drawing.Size(69, 20)
        Me.Label165.TabIndex = 16
        Me.Label165.Text = "Barcode"
        Me.Label165.Visible = False
        '
        'Label164
        '
        Me.Label164.AutoSize = True
        Me.Label164.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label164.Location = New System.Drawing.Point(23, 77)
        Me.Label164.Name = "Label164"
        Me.Label164.Size = New System.Drawing.Size(75, 20)
        Me.Label164.TabIndex = 15
        Me.Label164.Text = "QR Code"
        '
        'Label163
        '
        Me.Label163.AutoSize = True
        Me.Label163.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label163.Location = New System.Drawing.Point(23, 164)
        Me.Label163.Name = "Label163"
        Me.Label163.Size = New System.Drawing.Size(177, 20)
        Me.Label163.TabIndex = 14
        Me.Label163.Text = "Plant Code + DateCode"
        '
        'txt_PlanCode_DateCode
        '
        Me.txt_PlanCode_DateCode.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_PlanCode_DateCode.Location = New System.Drawing.Point(204, 164)
        Me.txt_PlanCode_DateCode.Name = "txt_PlanCode_DateCode"
        Me.txt_PlanCode_DateCode.Size = New System.Drawing.Size(146, 26)
        Me.txt_PlanCode_DateCode.TabIndex = 13
        '
        'txt_QRCode
        '
        Me.txt_QRCode.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_QRCode.Location = New System.Drawing.Point(204, 77)
        Me.txt_QRCode.Name = "txt_QRCode"
        Me.txt_QRCode.Size = New System.Drawing.Size(378, 26)
        Me.txt_QRCode.TabIndex = 12
        '
        'btn_Cek
        '
        Me.btn_Cek.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_Cek.Location = New System.Drawing.Point(447, 152)
        Me.btn_Cek.Name = "btn_Cek"
        Me.btn_Cek.Size = New System.Drawing.Size(135, 50)
        Me.btn_Cek.TabIndex = 19
        Me.btn_Cek.Text = "OK"
        Me.btn_Cek.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 21)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(586, 20)
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "Please  Scan QRcode or Barcode or Type in Plan Code + Date Code"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(200, 193)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(152, 20)
        Me.Label2.TabIndex = 16
        Me.Label2.Text = "Example :  AF20151"
        '
        'Input_BoX
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(610, 230)
        Me.Controls.Add(Me.btn_Cek)
        Me.Controls.Add(Me.Combo_PlanCode)
        Me.Controls.Add(Me.Txt_Barcode)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label165)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label164)
        Me.Controls.Add(Me.Label163)
        Me.Controls.Add(Me.txt_PlanCode_DateCode)
        Me.Controls.Add(Me.txt_QRCode)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "Input_BoX"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Quality Issue Cheking"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Combo_PlanCode As ComboBox
    Friend WithEvents Txt_Barcode As TextBox
    Friend WithEvents Label165 As Label
    Friend WithEvents Label164 As Label
    Friend WithEvents Label163 As Label
    Friend WithEvents txt_PlanCode_DateCode As TextBox
    Friend WithEvents txt_QRCode As TextBox
    Friend WithEvents btn_Cek As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
End Class
