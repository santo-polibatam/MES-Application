<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Export_to_Excel
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
        Me.components = New System.ComponentModel.Container()
        Me.NSXMasterdataBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.SGRAC_MESDataSet = New LAC.SGRAC_MESDataSet()
        Me.OpenOrdersBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.OpenOrdersTableAdapter = New LAC.SGRAC_MESDataSetTableAdapters.openOrdersTableAdapter()
        Me.BtnExport = New System.Windows.Forms.Button()
        Me.NSXMasterdataTableAdapter = New LAC.SGRAC_MESDataSetTableAdapters.NSXMasterdataTableAdapter()
        Me.NSXMasterdataBindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.IDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.MaterialDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ComponentsBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.ComponentsTableAdapter = New LAC.SGRAC_MESDataSetTableAdapters.ComponentsTableAdapter()
        Me.OpenOrdersBindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        CType(Me.NSXMasterdataBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SGRAC_MESDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.OpenOrdersBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NSXMasterdataBindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ComponentsBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.OpenOrdersBindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'NSXMasterdataBindingSource
        '
        Me.NSXMasterdataBindingSource.DataMember = "NSXMasterdata"
        Me.NSXMasterdataBindingSource.DataSource = Me.SGRAC_MESDataSet
        '
        'SGRAC_MESDataSet
        '
        Me.SGRAC_MESDataSet.DataSetName = "SGRAC_MESDataSet"
        Me.SGRAC_MESDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'OpenOrdersBindingSource
        '
        Me.OpenOrdersBindingSource.DataMember = "openOrders"
        Me.OpenOrdersBindingSource.DataSource = Me.SGRAC_MESDataSet
        '
        'OpenOrdersTableAdapter
        '
        Me.OpenOrdersTableAdapter.ClearBeforeFill = True
        '
        'BtnExport
        '
        Me.BtnExport.Location = New System.Drawing.Point(427, 264)
        Me.BtnExport.Name = "BtnExport"
        Me.BtnExport.Size = New System.Drawing.Size(75, 23)
        Me.BtnExport.TabIndex = 1
        Me.BtnExport.Text = "Export"
        Me.BtnExport.UseVisualStyleBackColor = True
        '
        'NSXMasterdataTableAdapter
        '
        Me.NSXMasterdataTableAdapter.ClearBeforeFill = True
        '
        'NSXMasterdataBindingSource1
        '
        Me.NSXMasterdataBindingSource1.DataMember = "NSXMasterdata"
        Me.NSXMasterdataBindingSource1.DataSource = Me.SGRAC_MESDataSet
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToOrderColumns = True
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.IDDataGridViewTextBoxColumn, Me.MaterialDataGridViewTextBoxColumn})
        Me.DataGridView1.DataSource = Me.OpenOrdersBindingSource1
        Me.DataGridView1.Location = New System.Drawing.Point(98, 53)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(880, 150)
        Me.DataGridView1.TabIndex = 2
        '
        'IDDataGridViewTextBoxColumn
        '
        Me.IDDataGridViewTextBoxColumn.DataPropertyName = "ID"
        Me.IDDataGridViewTextBoxColumn.HeaderText = "ID"
        Me.IDDataGridViewTextBoxColumn.Name = "IDDataGridViewTextBoxColumn"
        Me.IDDataGridViewTextBoxColumn.ReadOnly = True
        '
        'MaterialDataGridViewTextBoxColumn
        '
        Me.MaterialDataGridViewTextBoxColumn.DataPropertyName = "Material"
        Me.MaterialDataGridViewTextBoxColumn.HeaderText = "Material"
        Me.MaterialDataGridViewTextBoxColumn.Name = "MaterialDataGridViewTextBoxColumn"
        '
        'ComponentsBindingSource
        '
        Me.ComponentsBindingSource.DataMember = "Components"
        Me.ComponentsBindingSource.DataSource = Me.SGRAC_MESDataSet
        '
        'ComponentsTableAdapter
        '
        Me.ComponentsTableAdapter.ClearBeforeFill = True
        '
        'OpenOrdersBindingSource1
        '
        Me.OpenOrdersBindingSource1.DataMember = "openOrders"
        Me.OpenOrdersBindingSource1.DataSource = Me.SGRAC_MESDataSet
        '
        'Export_to_Excel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1034, 540)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.BtnExport)
        Me.Name = "Export_to_Excel"
        Me.Text = "Export_to_Excel"
        CType(Me.NSXMasterdataBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SGRAC_MESDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.OpenOrdersBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NSXMasterdataBindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ComponentsBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.OpenOrdersBindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SGRAC_MESDataSet As SGRAC_MESDataSet
    Friend WithEvents OpenOrdersBindingSource As BindingSource
    Friend WithEvents OpenOrdersTableAdapter As SGRAC_MESDataSetTableAdapters.openOrdersTableAdapter
    Friend WithEvents BtnExport As Button
    Friend WithEvents NSXMasterdataBindingSource As BindingSource
    Friend WithEvents NSXMasterdataTableAdapter As SGRAC_MESDataSetTableAdapters.NSXMasterdataTableAdapter
    Friend WithEvents NSXMasterdataBindingSource1 As BindingSource
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents IDDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents MaterialDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents ComponentsBindingSource As BindingSource
    Friend WithEvents ComponentsTableAdapter As SGRAC_MESDataSetTableAdapters.ComponentsTableAdapter
    Friend WithEvents OpenOrdersBindingSource1 As BindingSource
End Class
