Imports System.Data.SqlClient


Public Class Form1

    Public publicQuery As String
    Public Shared Sub open_form(query)
        Form1.publicQuery = query
        Call Main.koneksi_db()
        Try
            Dim sc As New SqlCommand(query, Main.koneksi)
            Dim adapter As New SqlDataAdapter(sc)
            Dim ds As New DataSet

            adapter.Fill(ds)
            Form1.DataGridView1.DataSource = ds.Tables(0)
            Form1.DataGridView1.Rows(0).Selected = True

            adapter.UpdateCommand = New SqlCommandBuilder(adapter).GetUpdateCommand
            adapter.Update(ds)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Shared Sub open_form_import_pp(imp As DataSet)
        Try
            Form1.DataGridView1.ColumnCount = 8
            Form1.DataGridView1.Columns(0).Name = "No"
            Form1.DataGridView1.Columns(1).Name = "Order Number"
            Form1.DataGridView1.Columns(2).Name = "Material"
            Form1.DataGridView1.Columns(3).Name = "Description"
            Form1.DataGridView1.Columns(4).Name = "Reqmts Qty"
            Form1.DataGridView1.Columns(5).Name = "Stor Loc"
            Form1.DataGridView1.Columns(6).Name = "Status"
            Form1.DataGridView1.Columns(7).Name = "PeggedReq"

            For r = 5 To imp.Tables(0).Rows.Count - 1
                Dim row As String() = New String() {r - 4, imp.Tables(0).Rows(r).Item(1).ToString(), imp.Tables(0).Rows(r).Item(2).ToString(), imp.Tables(0).Rows(r).Item(3).ToString(), imp.Tables(0).Rows(r).Item(4).ToString(), imp.Tables(0).Rows(r).Item(5).ToString(), imp.Tables(0).Rows(r).Item(6).ToString(), imp.Tables(0).Rows(r).Item(7).ToString()}
                Form1.DataGridView1.Rows.Add(row)
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Shared Sub open_form_import_pplist(imp As DataSet)
        Try

            Form1.DataGridView1.ColumnCount = 39
            Form1.DataGridView1.Columns(0).Name = "No"
            Form1.DataGridView1.Columns(1).Name = "Create On"
            Form1.DataGridView1.Columns(2).Name = "Entered By"
            Form1.DataGridView1.Columns(3).Name = "Type"
            Form1.DataGridView1.Columns(4).Name = "Order"
            Form1.DataGridView1.Columns(5).Name = "Start time"
            Form1.DataGridView1.Columns(6).Name = "Basic fin"
            Form1.DataGridView1.Columns(7).Name = "Committed"

            Form1.DataGridView1.Columns(8).Name = "SchedStart"
            Form1.DataGridView1.Columns(9).Name = "Finish tme"
            Form1.DataGridView1.Columns(10).Name = "PrCtr"
            Form1.DataGridView1.Columns(11).Name = "Material"
            Form1.DataGridView1.Columns(12).Name = "Scheduled finish"
            Form1.DataGridView1.Columns(13).Name = "Scheduled start"
            Form1.DataGridView1.Columns(14).Name = "Basic start date"
            Form1.DataGridView1.Columns(15).Name = "Basic finish date"

            Form1.DataGridView1.Columns(16).Name = "Item quantity"
            Form1.DataGridView1.Columns(17).Name = "OUM"
            Form1.DataGridView1.Columns(18).Name = "Material Description"
            Form1.DataGridView1.Columns(19).Name = "Description"
            Form1.DataGridView1.Columns(20).Name = "CDI"
            Form1.DataGridView1.Columns(21).Name = "Req dlv dt"
            Form1.DataGridView1.Columns(22).Name = "Purchase order number"
            Form1.DataGridView1.Columns(23).Name = "POitem"

            Form1.DataGridView1.Columns(24).Name = "GI time"
            Form1.DataGridView1.Columns(25).Name = "Stag time"
            Form1.DataGridView1.Columns(26).Name = "GI date"
            Form1.DataGridView1.Columns(27).Name = "Mat av dt"
            Form1.DataGridView1.Columns(28).Name = "Confirmed qty"
            Form1.DataGridView1.Columns(29).Name = "SU"
            Form1.DataGridView1.Columns(30).Name = "Customer"
            Form1.DataGridView1.Columns(31).Name = "City"
            Form1.DataGridView1.Columns(32).Name = "Name 2"

            Form1.DataGridView1.Columns(33).Name = "Name 1"
            Form1.DataGridView1.Columns(34).Name = "Item"
            Form1.DataGridView1.Columns(35).Name = "So No"
            Form1.DataGridView1.Columns(36).Name = "Customer Material Number"
            Form1.DataGridView1.Columns(37).Name = "Status"
            Form1.DataGridView1.Columns(38).Name = "Stat"

        Dim field0 As String
        Dim field4 As String
        Dim field5 As String
        Dim field7 As String
        Dim field8 As String
        Dim field11 As String
        Dim field12 As String
        Dim field13 As String
        Dim field14 As String
        Dim field20 As String
        Dim field23 As String
        Dim field24 As String
        Dim field25 As String
        Dim field26 As String

        For r = 1 To imp.Tables(0).Rows.Count - 1
            If String.IsNullOrEmpty(imp.Tables(0).Rows(r).Item(0).ToString()) = True Then
                field0 = ""
            Else
                    field0 = New DateTime().FromOADate(imp.Tables(0).Rows(r).Item(0).ToString())
                End If

            If String.IsNullOrEmpty(imp.Tables(0).Rows(r).Item(4).ToString()) = True Then
                field4 = ""
            Else
                field4 = New DateTime().FromOADate(imp.Tables(0).Rows(r).Item(4).ToString())
            End If

            If String.IsNullOrEmpty(imp.Tables(0).Rows(r).Item(5).ToString()) = True Then
                field5 = ""
            Else
                field5 = New DateTime().FromOADate(imp.Tables(0).Rows(r).Item(5).ToString())
            End If

            If String.IsNullOrEmpty(imp.Tables(0).Rows(r).Item(7).ToString()) = True Then
                field7 = ""
            Else
                field7 = New DateTime().FromOADate(imp.Tables(0).Rows(r).Item(7).ToString())
            End If

            If String.IsNullOrEmpty(imp.Tables(0).Rows(r).Item(8).ToString()) = True Then
                field8 = ""
            Else
                field8 = New DateTime().FromOADate(imp.Tables(0).Rows(r).Item(8).ToString())
            End If

            If String.IsNullOrEmpty(imp.Tables(0).Rows(r).Item(11).ToString()) = True Then
                field11 = ""
            Else
                field11 = New DateTime().FromOADate(imp.Tables(0).Rows(r).Item(11).ToString())
            End If

            If String.IsNullOrEmpty(imp.Tables(0).Rows(r).Item(12).ToString()) = True Then
                field12 = ""
            Else
                field12 = New DateTime().FromOADate(imp.Tables(0).Rows(r).Item(12).ToString())
            End If

            If String.IsNullOrEmpty(imp.Tables(0).Rows(r).Item(13).ToString()) = True Then
                field13 = ""
            Else
                field13 = New DateTime().FromOADate(imp.Tables(0).Rows(r).Item(13).ToString())
            End If

            If String.IsNullOrEmpty(imp.Tables(0).Rows(r).Item(14).ToString()) = True Then
                field14 = ""
            Else
                field14 = New DateTime().FromOADate(imp.Tables(0).Rows(r).Item(14).ToString())
            End If

            If String.IsNullOrEmpty(imp.Tables(0).Rows(r).Item(20).ToString()) = True Then
                field20 = ""
            Else
                field20 = New DateTime().FromOADate(imp.Tables(0).Rows(r).Item(20).ToString())
            End If

            If String.IsNullOrEmpty(imp.Tables(0).Rows(r).Item(23).ToString()) = True Then
                field23 = ""
            Else
                field23 = New DateTime().FromOADate(imp.Tables(0).Rows(r).Item(23).ToString())
            End If

            If String.IsNullOrEmpty(imp.Tables(0).Rows(r).Item(24).ToString()) = True Then
                field24 = ""
            Else
                field24 = New DateTime().FromOADate(imp.Tables(0).Rows(r).Item(24).ToString())
            End If

            If String.IsNullOrEmpty(imp.Tables(0).Rows(r).Item(25).ToString()) = True Then
                field25 = ""
            Else
                field25 = New DateTime().FromOADate(imp.Tables(0).Rows(r).Item(25).ToString())
            End If

            If String.IsNullOrEmpty(imp.Tables(0).Rows(r).Item(26).ToString()) = True Then
                field26 = ""
            Else
                field26 = New DateTime().FromOADate(imp.Tables(0).Rows(r).Item(26).ToString())
            End If

            Dim row As String() = New String() {r,
                        field0,
                        imp.Tables(0).Rows(r).Item(1).ToString(),
                        imp.Tables(0).Rows(r).Item(2).ToString(),
                        imp.Tables(0).Rows(r).Item(3).ToString(),
                        field4,
                        field5,
                        imp.Tables(0).Rows(r).Item(6).ToString(),
                        field7,
                        field8,
                        imp.Tables(0).Rows(r).Item(9).ToString(),
                        imp.Tables(0).Rows(r).Item(10).ToString(),
                        field11,
                        field12,
                        field13,
                        field14,
                        imp.Tables(0).Rows(r).Item(15).ToString(),
                        imp.Tables(0).Rows(r).Item(16).ToString(),
                        imp.Tables(0).Rows(r).Item(17).ToString(),
                        imp.Tables(0).Rows(r).Item(18).ToString(),
                        imp.Tables(0).Rows(r).Item(19).ToString(),
                        field20,
                        imp.Tables(0).Rows(r).Item(21).ToString(),
                        imp.Tables(0).Rows(r).Item(22).ToString(),
                        field23,
                        field24,
                        field25,
                        field26,
                        imp.Tables(0).Rows(r).Item(27).ToString(),
                        imp.Tables(0).Rows(r).Item(28).ToString(),
                        imp.Tables(0).Rows(r).Item(29).ToString(),
                        imp.Tables(0).Rows(r).Item(30).ToString(),
                        imp.Tables(0).Rows(r).Item(31).ToString(),
                        imp.Tables(0).Rows(r).Item(32).ToString(),
                        imp.Tables(0).Rows(r).Item(33).ToString(),
                        imp.Tables(0).Rows(r).Item(34).ToString(),
                        imp.Tables(0).Rows(r).Item(35).ToString(),
                        imp.Tables(0).Rows(r).Item(36).ToString(),
                        imp.Tables(0).Rows(r).Item(37).ToString()
                    }
            Form1.DataGridView1.Rows.Add(row)
            Next
        Catch ex As Exception
        'MessageBox.Show("--asd" + ex.Message)
        End Try
    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        'Dim sc As New SqlCommand(publicQuery, Main.koneksi)
        'Dim adapter As New SqlDataAdapter(sc)
        'Dim ds As New DataSet
        'adapter.Update(ds)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'DataGridView1.AutoResizeColumns()
        TextBox1.Select()
    End Sub

    Private Sub DataGridView1_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs) Handles DataGridView1.DataBindingComplete
        Label1.Text = "Total Rows : " & DataGridView1.Rows.Count - 1
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim str As String = TextBox1.Text
        Try
            If Me.TextBox1.Text.Trim(" ") = " " Then
            Else
                For i As Integer = 0 To DataGridView1.Rows.Count - 1
                    For j As Integer = 0 To Me.DataGridView1.Rows(i).Cells.Count - 1
                        If DataGridView1.Item(j, i).Value.ToString().ToLower.StartsWith(str.ToLower) Or InStr(DataGridView1.Item(j, i).Value.ToString().ToLower, str.ToLower) Then
                            DataGridView1.Rows(i).Selected = True
                            DataGridView1.CurrentCell = DataGridView1.Rows(i).Cells(j)
                            Exit Sub
                        End If
                    Next
                Next i
            End If

        Catch abc As Exception
        End Try
        MsgBox("Data not found!")
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Dim str As String = TextBox1.Text
            Try
                If Me.TextBox1.Text.Trim(" ") = " " Then
                Else
                    For i As Integer = 0 To DataGridView1.Rows.Count - 1
                        For j As Integer = 0 To Me.DataGridView1.Rows(i).Cells.Count - 1
                            If DataGridView1.Item(j, i).Value.ToString().ToLower.StartsWith(str.ToLower) Or InStr(DataGridView1.Item(j, i).Value.ToString().ToLower, str.ToLower) Then
                                DataGridView1.Rows(i).Selected = True
                                DataGridView1.CurrentCell = DataGridView1.Rows(i).Cells(j)
                                Exit Sub
                            End If
                        Next
                    Next i
                End If

            Catch abc As Exception
            End Try
            MsgBox("Data not found!")
        End If
    End Sub

End Class