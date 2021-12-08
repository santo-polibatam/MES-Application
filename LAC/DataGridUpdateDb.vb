Imports System.Data.SqlClient

Public Class DataGridUpdateDb

    Public Property OKNEXT_Var As Integer

    Public publicQuery As String
    Private Sub txt_Search_db_TextChanged(sender As Object, e As KeyPressEventArgs) Handles txt_Search_db.KeyPress
        If e.KeyChar = Chr(13) Then
            reload()
            Dim str As String = txt_Search_db.Text
            Try
                If Me.txt_Search_db.Text.Trim(" ") = " " Then
                Else
                    For i As Integer = 0 To DataGridView1.Rows.Count - 1
                        For j As Integer = 0 To Me.DataGridView1.Rows(i).Cells.Count - 1
                            If DataGridView1.Item(j, i).Value.ToString().ToLower.StartsWith(str.ToLower) Or InStr(DataGridView1.Item(j, i).Value.ToString().ToLower, str.ToLower) Then
                                DataGridView1.Rows(i).Selected = True
                                DataGridView1.CurrentCell = DataGridView1.Rows(i).Cells(j)

                                OKNext.ShowDialog()

                                If OKNEXT_Var = 1 Then
                                    Exit Sub
                                End If

                                'Exit Sub
                            End If
                        Next
                    Next i
                End If

            Catch abc As Exception
            End Try
            MsgBox("Data not found!")
        End If
    End Sub
    Private Sub reload()
        Dim query As String = "select * from NewScanningComponent"
        Call Main.koneksi_db()
        Try
            Dim sc As New SqlCommand(query, Main.koneksi)
            Dim adapter As New SqlDataAdapter(sc)
            Dim ds As New DataSet

            adapter.Fill(ds)
            DataGridView1.DataSource = ds.Tables(0)
            DataGridView1.Rows(0).Selected = True

            'adapter.UpdateCommand = New SqlCommandBuilder(adapter).GetUpdateCommand
            'adapter.Update(ds)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub DataGridUpdateDb_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'DataGridUpdateDb.publicQuery = query
        reload()
    End Sub

    Private Sub Btn_update_Db_Click(sender As Object, e As EventArgs) Handles Btn_update_Db.Click

        If DataGridView1.Rows.Count > 1 Then
            Dim adapter As New SqlDataAdapter
            Dim query As String
            Dim numberOfData As Integer = DataGridView1.Rows.Count
            Call Main.koneksi_db()
            ToolStripProgressBar1.Value = 0

            For i As Integer = 0 To DataGridView1.Rows.Count - 1

                'query = "update SGRAC_MES.dbo.NewScanningComponent set Material='" & DataGridView1.Rows(i).Cells(0).Value & "',Description='" & DataGridView1.Rows(i).Cells(1).Value & "',Category='" & DataGridView1.Rows(i).Cells(2).Value & "',[Future Order]='" & DataGridView1.Rows(i).Cells(3).Value & "',[QR code]='" & DataGridView1.Rows(i).Cells(4).Value & "',Reference='" & DataGridView1.Rows(i).Cells(5).Value & "' where Material='" & DataGridView1.Rows(i).Cells(0).Value & "'"
                query = "update SGRAC_MES.dbo.NewScanningComponent set Material='" & DataGridView1.Rows(i).Cells(0).Value & "',Description='" & DataGridView1.Rows(i).Cells(1).Value & "',Category='" & DataGridView1.Rows(i).Cells(2).Value & "',[QR code]='" & DataGridView1.Rows(i).Cells(3).Value & "',Reference='" & DataGridView1.Rows(i).Cells(4).Value & "' where Material='" & DataGridView1.Rows(i).Cells(0).Value & "'"

                adapter = New SqlDataAdapter(query, Main.koneksi)
                adapter.SelectCommand.ExecuteNonQuery()

                Dim a As Integer = i / numberOfData * 100
                ToolStripProgressBar1.Value = a
                Application.DoEvents()
            Next
            MessageBox.Show("New Database updated successfully")
        End If

    End Sub

    Private Sub btn_refresh_Click(sender As Object, e As EventArgs) Handles btn_refresh.Click
        reload()
    End Sub

    Private Sub btn_Add_Click(sender As Object, e As EventArgs) Handles btn_Add.Click
        'Dim lastRows As Integer = DataGridView1.Rows.Count - 2
        'Dim data(5) As String
        'For i As Integer = 0 To 4
        '    data(i) = ""
        '    If Not IsDBNull(DataGridView1.Rows(lastRows).Cells(i).Value) Then
        '        data(i) = DataGridView1.Rows(lastRows).Cells(i).Value
        '    End If

        'Next

        'Dim query As String = "insert into SGRAC_MES.dbo.NewScanningComponent values(' ',' ',' ',' ',' ',' ')"
        'Dim query As String = "insert into SGRAC_MES.dbo.NewScanningComponent values('" & data(0) & "','" & data(1) & "','" & data(2) & "','" & data(3) & "','" & data(4) & "','" & data(5) & "')"
        Dim query As String = "insert into SGRAC_MES.dbo.NewScanningComponent values('" & txtMaterial.Text & "','" & txtxDescription.Text & "','" & txtCategory.Text & "','" & txtQRCode.Text & "','" & txtReferences.Text & "')"
        Dim adapter As New SqlDataAdapter
        Call Main.koneksi_db()
        adapter = New SqlDataAdapter(query, Main.koneksi)
        adapter.SelectCommand.ExecuteNonQuery()
        reload()

        Dim str As String = txtMaterial.Text

        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            For j As Integer = 0 To Me.DataGridView1.Rows(i).Cells.Count - 1
                If DataGridView1.Item(j, i).Value.ToString().ToLower.StartsWith(Str.ToLower) Or InStr(DataGridView1.Item(j, i).Value.ToString().ToLower, Str.ToLower) Then
                    DataGridView1.Rows(i).Selected = True
                    DataGridView1.CurrentCell = DataGridView1.Rows(i).Cells(j)


                    txtMaterial.Text = ""
                    txtxDescription.Text = ""
                    txtCategory.Text = ""
                    txtQRCode.Text = ""
                    txtReferences.Text = ""

                    Exit Sub

                End If
            Next
        Next i

        txtMaterial.Text = ""
        txtxDescription.Text = ""
        txtCategory.Text = ""
        txtQRCode.Text = ""
        txtReferences.Text = ""

        'DataGridView1.Rows(DataGridView1.Rows.Count - 2).Selected = True
        'DataGridView1.FirstDisplayedScrollingRowIndex = DataGridView1.Rows.Count - 1
    End Sub

    Private Sub btn_Save_Click(sender As Object, e As EventArgs) Handles btn_Delete.Click
        Dim Material As String = DataGridView1.SelectedRows(0).Cells(0).Value.ToString

        Dim result As DialogResult = MessageBox.Show("Are You Sure delete Mataerial : " & Material, "Delete Row", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Exit Sub
        ElseIf result = DialogResult.Yes Then
            Dim query As String = "delete from SGRAC_MES.dbo.NewScanningComponent where Material = '" & Material & "'"
            Dim adapter As New SqlDataAdapter
            Call Main.koneksi_db()
            adapter = New SqlDataAdapter(query, Main.koneksi)
            adapter.SelectCommand.ExecuteNonQuery()
            reload()
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Dim sql As String = "with deleteDup as(select *, ROW_NUMBER() over (partition by [order] order by id) as RowNumber from [SGRAC_MES].[dbo].[PPList]) delete from deleteDup where RowNumber > 1"
        Dim sql As String = "SELECT * FROM SGRAC_MES.dbo.NewScanningComponent
        ;WITH CTE AS(
            SELECT [Material],
            RN = ROW_NUMBER()OVER(PARTITION BY  [Material] ORDER BY [Material])
            FROM SGRAC_MES.dbo.NewScanningComponent
            )
        DELETE FROM CTE WHERE RN <> 1"
        Dim cmd As New SqlCommand(Sql, Main.koneksi)
        Dim count As Integer = cmd.ExecuteNonQuery()
        MsgBox("The Number Of Duplicate data has been executed : " & count & " Records")
        reload()
    End Sub

    Private Sub Cek_Duplicate()
        Dim query As String = "SELECT SGRAC_MES.dbo.NewScanningComponent.Material , COUNT(*)
                             FROM SGRAC_MES.dbo.NewScanningComponent
                            GROUP BY SGRAC_MES.dbo.NewScanningComponent.Material
                            HAVING COUNT(*) > 1"

        Call Main.koneksi_db()
        Try
            Dim sc As New SqlCommand(query, Main.koneksi)
            Dim adapter As New SqlDataAdapter(sc)
            Dim ds As New DataSet

            adapter.Fill(ds)
            DataGridView1.DataSource = ds.Tables(0)
            DataGridView1.Rows(0).Selected = True

            'adapter.UpdateCommand = New SqlCommandBuilder(adapter).GetUpdateCommand
            'adapter.Update(ds)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

End Class