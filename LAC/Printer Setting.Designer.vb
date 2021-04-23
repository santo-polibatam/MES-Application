<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Printer_Setting
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Printer_Setting))
        Me.DGV_Printer_Setting = New System.Windows.Forms.DataGridView()
        Me.IDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.WorkstationDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ReportTypeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PrinterDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PrinterTypeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DpiDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Field1DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PrinterTableBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.SGRAC_MESDataSet = New LAC.SGRAC_MESDataSet()
        Me.PrinterTableTableAdapter = New LAC.SGRAC_MESDataSetTableAdapters.printerTableTableAdapter()
        Me.FillByToolStrip = New System.Windows.Forms.ToolStrip()
        Me.FillByToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.Save = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripButton2 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        CType(Me.DGV_Printer_Setting, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PrinterTableBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SGRAC_MESDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.FillByToolStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'DGV_Printer_Setting
        '
        Me.DGV_Printer_Setting.AllowUserToOrderColumns = True
        Me.DGV_Printer_Setting.AutoGenerateColumns = False
        Me.DGV_Printer_Setting.BackgroundColor = System.Drawing.SystemColors.ButtonFace
        Me.DGV_Printer_Setting.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV_Printer_Setting.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.IDDataGridViewTextBoxColumn, Me.WorkstationDataGridViewTextBoxColumn, Me.ReportTypeDataGridViewTextBoxColumn, Me.PrinterDataGridViewTextBoxColumn, Me.PrinterTypeDataGridViewTextBoxColumn, Me.DpiDataGridViewTextBoxColumn, Me.Field1DataGridViewTextBoxColumn})
        Me.DGV_Printer_Setting.DataSource = Me.PrinterTableBindingSource
        Me.DGV_Printer_Setting.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DGV_Printer_Setting.Location = New System.Drawing.Point(0, 0)
        Me.DGV_Printer_Setting.Name = "DGV_Printer_Setting"
        Me.DGV_Printer_Setting.Size = New System.Drawing.Size(800, 691)
        Me.DGV_Printer_Setting.TabIndex = 0
        '
        'IDDataGridViewTextBoxColumn
        '
        Me.IDDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.IDDataGridViewTextBoxColumn.DataPropertyName = "ID"
        Me.IDDataGridViewTextBoxColumn.HeaderText = "ID"
        Me.IDDataGridViewTextBoxColumn.Name = "IDDataGridViewTextBoxColumn"
        Me.IDDataGridViewTextBoxColumn.ReadOnly = True
        Me.IDDataGridViewTextBoxColumn.Width = 43
        '
        'WorkstationDataGridViewTextBoxColumn
        '
        Me.WorkstationDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.WorkstationDataGridViewTextBoxColumn.DataPropertyName = "workstation"
        Me.WorkstationDataGridViewTextBoxColumn.HeaderText = "workstation"
        Me.WorkstationDataGridViewTextBoxColumn.Name = "WorkstationDataGridViewTextBoxColumn"
        Me.WorkstationDataGridViewTextBoxColumn.Width = 86
        '
        'ReportTypeDataGridViewTextBoxColumn
        '
        Me.ReportTypeDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.ReportTypeDataGridViewTextBoxColumn.DataPropertyName = "report type"
        Me.ReportTypeDataGridViewTextBoxColumn.HeaderText = "report type"
        Me.ReportTypeDataGridViewTextBoxColumn.Name = "ReportTypeDataGridViewTextBoxColumn"
        Me.ReportTypeDataGridViewTextBoxColumn.Width = 82
        '
        'PrinterDataGridViewTextBoxColumn
        '
        Me.PrinterDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.PrinterDataGridViewTextBoxColumn.DataPropertyName = "printer"
        Me.PrinterDataGridViewTextBoxColumn.HeaderText = "printer"
        Me.PrinterDataGridViewTextBoxColumn.Name = "PrinterDataGridViewTextBoxColumn"
        Me.PrinterDataGridViewTextBoxColumn.Width = 61
        '
        'PrinterTypeDataGridViewTextBoxColumn
        '
        Me.PrinterTypeDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.PrinterTypeDataGridViewTextBoxColumn.DataPropertyName = "printer type"
        Me.PrinterTypeDataGridViewTextBoxColumn.HeaderText = "printer type"
        Me.PrinterTypeDataGridViewTextBoxColumn.Name = "PrinterTypeDataGridViewTextBoxColumn"
        Me.PrinterTypeDataGridViewTextBoxColumn.Width = 84
        '
        'DpiDataGridViewTextBoxColumn
        '
        Me.DpiDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.DpiDataGridViewTextBoxColumn.DataPropertyName = "dpi"
        Me.DpiDataGridViewTextBoxColumn.HeaderText = "dpi"
        Me.DpiDataGridViewTextBoxColumn.Name = "DpiDataGridViewTextBoxColumn"
        Me.DpiDataGridViewTextBoxColumn.Width = 46
        '
        'Field1DataGridViewTextBoxColumn
        '
        Me.Field1DataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Field1DataGridViewTextBoxColumn.DataPropertyName = "Field1"
        Me.Field1DataGridViewTextBoxColumn.HeaderText = "Field1"
        Me.Field1DataGridViewTextBoxColumn.Name = "Field1DataGridViewTextBoxColumn"
        Me.Field1DataGridViewTextBoxColumn.Width = 60
        '
        'PrinterTableBindingSource
        '
        Me.PrinterTableBindingSource.DataMember = "printerTable"
        Me.PrinterTableBindingSource.DataSource = Me.SGRAC_MESDataSet
        '
        'SGRAC_MESDataSet
        '
        Me.SGRAC_MESDataSet.DataSetName = "SGRAC_MESDataSet"
        Me.SGRAC_MESDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'PrinterTableTableAdapter
        '
        Me.PrinterTableTableAdapter.ClearBeforeFill = True
        '
        'FillByToolStrip
        '
        Me.FillByToolStrip.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.FillByToolStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FillByToolStripButton, Me.ToolStripSeparator1, Me.Save, Me.ToolStripSeparator2, Me.ToolStripButton1, Me.ToolStripSeparator3, Me.ToolStripButton2, Me.ToolStripSeparator4})
        Me.FillByToolStrip.Location = New System.Drawing.Point(0, 666)
        Me.FillByToolStrip.Name = "FillByToolStrip"
        Me.FillByToolStrip.Size = New System.Drawing.Size(800, 25)
        Me.FillByToolStrip.TabIndex = 3
        Me.FillByToolStrip.Text = "FillByToolStrip"
        '
        'FillByToolStripButton
        '
        Me.FillByToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.FillByToolStripButton.Name = "FillByToolStripButton"
        Me.FillByToolStripButton.Size = New System.Drawing.Size(50, 22)
        Me.FillByToolStripButton.Text = "Refresh"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'Save
        '
        Me.Save.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.Save.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.Save.Name = "Save"
        Me.Save.Size = New System.Drawing.Size(35, 22)
        Me.Save.Text = "Save"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 25)
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton1.Image = CType(resources.GetObject("ToolStripButton1.Image"), System.Drawing.Image)
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(33, 22)
        Me.ToolStripButton1.Text = "Add"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(6, 25)
        '
        'ToolStripButton2
        '
        Me.ToolStripButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton2.Image = CType(resources.GetObject("ToolStripButton2.Image"), System.Drawing.Image)
        Me.ToolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton2.Name = "ToolStripButton2"
        Me.ToolStripButton2.Size = New System.Drawing.Size(44, 22)
        Me.ToolStripButton2.Text = "Delete"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(6, 25)
        '
        'Printer_Setting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 691)
        Me.Controls.Add(Me.FillByToolStrip)
        Me.Controls.Add(Me.DGV_Printer_Setting)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Printer_Setting"
        Me.Text = "Printer_Setting"
        CType(Me.DGV_Printer_Setting, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PrinterTableBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SGRAC_MESDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.FillByToolStrip.ResumeLayout(False)
        Me.FillByToolStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DGV_Printer_Setting As DataGridView
    Friend WithEvents SGRAC_MESDataSet As SGRAC_MESDataSet
    Friend WithEvents PrinterTableBindingSource As BindingSource
    Friend WithEvents PrinterTableTableAdapter As SGRAC_MESDataSetTableAdapters.printerTableTableAdapter
    Friend WithEvents FillByToolStrip As ToolStrip
    Friend WithEvents FillByToolStripButton As ToolStripButton
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents Save As ToolStripButton
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents ToolStripButton1 As ToolStripButton
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents ToolStripButton2 As ToolStripButton
    Friend WithEvents ToolStripSeparator4 As ToolStripSeparator
    Friend WithEvents IDDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents WorkstationDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents ReportTypeDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents PrinterDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents PrinterTypeDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents DpiDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Field1DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
End Class
