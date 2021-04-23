Imports System.Data.SqlClient

Public Class Printer_Setting
    Dim rowIndex As Integer

    Dim _ID As Integer
    Dim _workstation As String
    Dim _report_type As String
    Dim _printer As String
    Dim _printer_type As String
    Dim _dpi As Integer

    Dim errorduringsaving As Integer = 0

    Private Sub Printer_Setting_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'SGRAC_MESDataSet.printerTable' table. You can move, or remove it, as needed.
        Me.PrinterTableTableAdapter.Fill(Me.SGRAC_MESDataSet.printerTable)

    End Sub

    Private Sub FillByToolStripButton_Click(sender As Object, e As EventArgs) Handles FillByToolStripButton.Click
        Try
            Me.PrinterTableTableAdapter.FillBy(Me.SGRAC_MESDataSet.printerTable)
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)

        End Try

    End Sub

    Private Sub FillByToolStripButton1_Click(sender As Object, e As EventArgs)
        Try
            Me.PrinterTableTableAdapter.FillBy(Me.SGRAC_MESDataSet.printerTable)
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub Save_Click(sender As Object, e As EventArgs) Handles Save.Click
        Try
            errorduringsaving = 0
            Call Main.koneksi_db()
            Dim sql As String = "UPDATE [dbo].[printerTable] SET [workstation] = '" &
                _workstation & "', [report type] = '" & _report_type & "', [printer] = '" &
                _printer & "', [printer type] = '" & _printer_type & "', [dpi] = " & _dpi &
                ", [Field1] = NULL WHERE        [ID] = " & _ID
            'MsgBox(sql)
            Dim adapterSimpan As New SqlCommand(sql, Main.koneksi)
            adapterSimpan.ExecuteNonQuery()
            load_DB_DGV()
            If errorduringsaving = 0 Then MessageBox.Show("Data Saved!",
            "Saving Proccess",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information,
            MessageBoxDefaultButton.Button1)
        Catch ex As Exception
            If errorduringsaving = 0 Then
                MessageBox.Show("Error During Saving !" & Chr(13) & Chr(13) & "Make Sure you press Enter and click saved Button.",
                "Important Note",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1)

                errorduringsaving = 1
            End If
        End Try

    End Sub

    Private Sub DGV_Printer_Setting_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DGV_Printer_Setting.CellEndEdit
        Try
            Dim row As DataGridViewRow
            row = Me.DGV_Printer_Setting.Rows(e.RowIndex)

            _ID = row.Cells(0).Value.ToString
            _workstation = row.Cells(1).Value.ToString
            _report_type = row.Cells(2).Value.ToString
            _printer = row.Cells(3).Value.ToString
            _printer_type = row.Cells(4).Value.ToString
            If Not String.IsNullOrEmpty(row.Cells(5).Value.ToString) Then _dpi = row.Cells(5).Value.ToString

            'MsgBox(_ID & _workstation)
        Catch ex As Exception
            'MsgBox(ex.Message)
            If errorduringsaving = 0 Then
                errorduringsaving = 1
                MessageBox.Show("Error During Saving !" & Chr(13) & Chr(13) & "Make Sure you press Enter and click saved Button.",
                "Important Note",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1)

            End If
        End Try
    End Sub

    Private Sub DGV_Printer_Setting_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DGV_Printer_Setting.CellMouseClick
        Try
            Dim row As DataGridViewRow
            row = Me.DGV_Printer_Setting.Rows(e.RowIndex)

            _ID = row.Cells(0).Value.ToString
            _workstation = row.Cells(1).Value.ToString
            _report_type = row.Cells(2).Value.ToString
            _printer = row.Cells(3).Value.ToString
            _printer_type = row.Cells(4).Value.ToString
            If Not String.IsNullOrEmpty(row.Cells(5).Value.ToString) Then _dpi = row.Cells(5).Value.ToString

            'MsgBox(_ID & _workstation)
        Catch ex As Exception
            'MsgBox(ex.Message)
            'If errorduringsaving = 0 Then
            '    MessageBox.Show("Error During Saving !" & Chr(13) & Chr(13) & "Make Sure you press Enter and click saved Button.",
            '    "Important Note",
            '    MessageBoxButtons.OK,
            '    MessageBoxIcon.Error,
            '    MessageBoxDefaultButton.Button1)

            '    errorduringsaving = 1
            'End If
        End Try
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Call Main.koneksi_db()
        Dim sql As String = "DELETE FROM [dbo].[printerTable] WHERE [id]=" & _ID
        Dim adapterSimpan As New SqlCommand(sql, Main.koneksi)
        Try
            adapterSimpan.ExecuteNonQuery()
            load_DB_DGV()
            MsgBox("Data Deleted")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Dim sql As String = "INSERT INTO [dbo].[printerTable]
           ([workstation]
           ,[report type]
           ,[printer]
           ,[printer type]
           ,[dpi]
           ,[Field1])
     VALUES
           (null,
		   null,
           null,
           null,
           null,
           null)"
        Dim adapterSimpan As New SqlCommand(sql, Main.koneksi)
        Try
            adapterSimpan.ExecuteNonQuery()
            load_DB_DGV()
            'MsgBox("Data Added")
            DGV_Printer_Setting.CurrentCell = DGV_Printer_Setting(1, DGV_Printer_Setting.Rows.Count - 2)
            DGV_Printer_Setting.BeginEdit(True)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    Private Sub load_DB_DGV()
        Try
            Me.PrinterTableTableAdapter.FillBy(Me.SGRAC_MESDataSet.printerTable)
        Catch ex As System.Exception
            'System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Printer_Setting_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Main.koneksi.Close()
    End Sub

    Private Sub DGV_Printer_Setting_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DGV_Printer_Setting.DataError
        If errorduringsaving = 0 Then
            MessageBox.Show("Error During Saving !" & Chr(13) & Chr(13) & "Make Sure you press Enter and click saved Button.",
            "Important Note",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error,
            MessageBoxDefaultButton.Button1)

            errorduringsaving = 1
        End If

    End Sub
End Class