Imports System.Data.SqlClient
Public Class Form2
    Dim rowIndex As Integer
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles Me.Load
        Button2.Enabled = False
        Call Main.koneksi_db()
        Dim imp As New DataSet
        Dim sql As String = "select * from PPList"
        Dim adapter As New SqlDataAdapter(sql, Main.koneksi)
        adapter.Fill(imp)
        'DataGridView1.DataSource = ds.Tables(0).Rows(1)

        Dim field20 As String

        DataGridView1.ColumnCount = 38
        DataGridView1.Columns(0).Name = "Create On"
        DataGridView1.Columns(1).Name = "Entered By"
        DataGridView1.Columns(2).Name = "Type"
        DataGridView1.Columns(3).Name = "Order"
        DataGridView1.Columns(4).Name = "Start time"
        DataGridView1.Columns(5).Name = "Basic fin"
        DataGridView1.Columns(6).Name = "Committed"

        DataGridView1.Columns(7).Name = "SchedStart"
        DataGridView1.Columns(8).Name = "Finish tme"
        DataGridView1.Columns(9).Name = "PrCtr"
        DataGridView1.Columns(10).Name = "Material"
        DataGridView1.Columns(11).Name = "Scheduled finish"
        DataGridView1.Columns(12).Name = "Scheduled start"
        DataGridView1.Columns(13).Name = "Basic start date"
        DataGridView1.Columns(14).Name = "Basic finish date"

        DataGridView1.Columns(15).Name = "Item quantity"
        DataGridView1.Columns(16).Name = "OUM"
        DataGridView1.Columns(17).Name = "Material Description"
        DataGridView1.Columns(18).Name = "Description"
        DataGridView1.Columns(19).Name = "CDI"
        DataGridView1.Columns(20).Name = "Req dlv dt"
        DataGridView1.Columns(21).Name = "Purchase order number"
        DataGridView1.Columns(22).Name = "POitem"

        DataGridView1.Columns(23).Name = "GI time"
        DataGridView1.Columns(24).Name = "Stag time"
        DataGridView1.Columns(25).Name = "GI date"
        DataGridView1.Columns(26).Name = "Mat av dt"
        DataGridView1.Columns(27).Name = "Confirmed qty"
        DataGridView1.Columns(28).Name = "SU"
        DataGridView1.Columns(29).Name = "Customer"
        DataGridView1.Columns(30).Name = "City"
        DataGridView1.Columns(31).Name = "Name 2"

        DataGridView1.Columns(32).Name = "Name 1"
        DataGridView1.Columns(33).Name = "Item"
        DataGridView1.Columns(34).Name = "So No"
        DataGridView1.Columns(35).Name = "Customer Material Number"
        DataGridView1.Columns(36).Name = "Status"
        DataGridView1.Columns(37).Name = "Stat"

        For r = 1 To imp.Tables(0).Rows.Count - 1
            If String.IsNullOrEmpty(imp.Tables(0).Rows(r).Item(21).ToString()) = True Then
                field20 = ""
            Else
                field20 = imp.Tables(0).Rows(r).Item(21).ToString()
            End If
            Dim row As String() = New String() {
                    imp.Tables(0).Rows(r).Item(1).ToString(),
                    imp.Tables(0).Rows(r).Item(2).ToString(),
                    imp.Tables(0).Rows(r).Item(3).ToString(),
                    imp.Tables(0).Rows(r).Item(4).ToString(),
                    imp.Tables(0).Rows(r).Item(5).ToString(),
                    imp.Tables(0).Rows(r).Item(6).ToString(),
                    imp.Tables(0).Rows(r).Item(7).ToString(),
                    imp.Tables(0).Rows(r).Item(8).ToString(),
                    imp.Tables(0).Rows(r).Item(9).ToString(),
                    imp.Tables(0).Rows(r).Item(10).ToString(),
                    imp.Tables(0).Rows(r).Item(11).ToString(),
                    imp.Tables(0).Rows(r).Item(12).ToString(),
                    imp.Tables(0).Rows(r).Item(13).ToString(),
                    imp.Tables(0).Rows(r).Item(14).ToString(),
                    imp.Tables(0).Rows(r).Item(15).ToString(),
                    imp.Tables(0).Rows(r).Item(16).ToString(),
                    imp.Tables(0).Rows(r).Item(17).ToString(),
                    imp.Tables(0).Rows(r).Item(18).ToString(),
                    imp.Tables(0).Rows(r).Item(19).ToString(),
                    imp.Tables(0).Rows(r).Item(20).ToString(),
                    field20,
                    imp.Tables(0).Rows(r).Item(22).ToString(),
                    imp.Tables(0).Rows(r).Item(23).ToString(),
                    imp.Tables(0).Rows(r).Item(24).ToString(),
                    imp.Tables(0).Rows(r).Item(25).ToString(),
                    imp.Tables(0).Rows(r).Item(26).ToString(),
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
                    imp.Tables(0).Rows(r).Item(37).ToString(),
                    imp.Tables(0).Rows(r).Item(38).ToString()
                }
            DataGridView1.Rows.Add(row)
        Next
    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        Button2.Enabled = True
        rowIndex = e.RowIndex
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Dim sql As String = "insert into [pplist] ([created on],[entered by],[type],[order],[start time],[basic fin],[committed],[schedstart],[finish tme],[prctr],[material],[scheduled finish] ,[scheduled start],[basic start date],[basic finish date],[item quantity],[oum],[material description],[description],[cdi],[req dlv dt],[purchase order number],[poitem],[gi time],[stag  time],[gi date],[mat av dt],[confirmed qty],[su],[customer],[city],[name 2],[name 1],[item],[so no],[customer material number],[status],[stat]) values(@create,@enter,@type,@order,@start,@basic,@com,@sched,@finish,@prctr,@material,@sched_finish,@sched_start,@basic_finish,@basic_start,@item_qty,@oum,@material_desc,@desc,@cdi,@req,@purchase,@po,@gi_ti,@stag,@gi_da,@mat,@confirm,@su,@customer,@city,@name2,@name1,@item,@so,@cusmat,@status,@stat)"
            Dim adapterSimpan As New SqlDataAdapter(sql, Main.koneksi)
            If String.IsNullOrEmpty(DataGridView1.Rows(rowIndex).Cells(0).ToString()) Then
                adapterSimpan.SelectCommand.Parameters.AddWithValue("@create", "00/00/0000")
            Else
                adapterSimpan.SelectCommand.Parameters.AddWithValue("@create", DataGridView1.Rows(rowIndex).Cells(0).ToString())
            End If
            'adapterSimpan.SelectCommand.Parameters.AddWithValue("@create", Convert.ToDateTime(DataGridView1.Rows(rowIndex).Cells(0).ToString()))
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@enter", DataGridView1.Rows(rowIndex).Cells(1).ToString())
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@type", DataGridView1.Rows(rowIndex).Cells(2).ToString())
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@order", DataGridView1.Rows(rowIndex).Cells(3).ToString())
            If String.IsNullOrEmpty(DataGridView1.Rows(rowIndex).Cells(4).ToString()) Then
                adapterSimpan.SelectCommand.Parameters.AddWithValue("@start", "00/00/0000")
            Else
                adapterSimpan.SelectCommand.Parameters.AddWithValue("@start", DataGridView1.Rows(rowIndex).Cells(4).ToString())
            End If
            If String.IsNullOrEmpty(DataGridView1.Rows(rowIndex).Cells(5).ToString()) Then
                adapterSimpan.SelectCommand.Parameters.AddWithValue("@basic", "00/00/0000")
            Else
                adapterSimpan.SelectCommand.Parameters.AddWithValue("@basic", DataGridView1.Rows(rowIndex).Cells(5).ToString())
            End If
            'adapterSimpan.SelectCommand.Parameters.AddWithValue("@start", Convert.ToDateTime(DataGridView1.Rows(rowIndex).Cells(4).ToString()))
            'adapterSimpan.SelectCommand.Parameters.AddWithValue("@basic", Convert.ToDateTime(DataGridView1.Rows(rowIndex).Cells(5).ToString()))
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@com", DataGridView1.Rows(rowIndex).Cells(6).ToString())
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@sched", Convert.ToDateTime(DataGridView1.Rows(rowIndex).Cells(7).ToString()))
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@finish", Convert.ToDateTime(DataGridView1.Rows(rowIndex).Cells(8).ToString()))
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@prctr", DataGridView1.Rows(rowIndex).Cells(9))
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@material", DataGridView1.Rows(rowIndex).Cells(10).ToString())
            If String.IsNullOrEmpty(DataGridView1.Rows(rowIndex).Cells(11).ToString()) Then
                adapterSimpan.SelectCommand.Parameters.AddWithValue("@sched_finish", "00/00/0000")
            Else
                adapterSimpan.SelectCommand.Parameters.AddWithValue("@sched_finish", DataGridView1.Rows(rowIndex).Cells(11).ToString())
            End If
            If String.IsNullOrEmpty(DataGridView1.Rows(rowIndex).Cells(12).ToString()) Then
                adapterSimpan.SelectCommand.Parameters.AddWithValue("@sched_start", "00/00/0000")
            Else
                adapterSimpan.SelectCommand.Parameters.AddWithValue("@sched_start", DataGridView1.Rows(rowIndex).Cells(12).ToString())
            End If
            If String.IsNullOrEmpty(DataGridView1.Rows(rowIndex).Cells(13).ToString()) Then
                adapterSimpan.SelectCommand.Parameters.AddWithValue("@basic_start", "00/00/0000")
            Else
                adapterSimpan.SelectCommand.Parameters.AddWithValue("@basic_start", DataGridView1.Rows(rowIndex).Cells(13).ToString())
            End If
            If String.IsNullOrEmpty(DataGridView1.Rows(rowIndex).Cells(14).ToString()) Then
                adapterSimpan.SelectCommand.Parameters.AddWithValue("@basic_finish", "00/00/0000")
            Else
                adapterSimpan.SelectCommand.Parameters.AddWithValue("@basic_finish", DataGridView1.Rows(rowIndex).Cells(14).ToString())
            End If
            'adapterSimpan.SelectCommand.Parameters.AddWithValue("@sched_finish", Convert.ToDateTime(DataGridView1.Rows(rowIndex).Cells(11).ToString()))
            'adapterSimpan.SelectCommand.Parameters.AddWithValue("@sched_start", Convert.ToDateTime(DataGridView1.Rows(rowIndex).Cells(12).ToString()))
            'adapterSimpan.SelectCommand.Parameters.AddWithValue("@basic_start", Convert.ToDateTime(DataGridView1.Rows(rowIndex).Cells(13).ToString()))
            'adapterSimpan.SelectCommand.Parameters.AddWithValue("@basic_finish", Convert.ToDateTime(DataGridView1.Rows(rowIndex).Cells(14).ToString()))
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@item_qty", DataGridView1.Rows(rowIndex).Cells(15))
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@oum", DataGridView1.Rows(rowIndex).Cells(16).ToString())
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@material_desc", DataGridView1.Rows(rowIndex).Cells(17).ToString())
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@desc", DataGridView1.Rows(rowIndex).Cells(18).ToString())
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@cdi", DataGridView1.Rows(rowIndex).Cells(19).ToString())
            If String.IsNullOrEmpty(DataGridView1.Rows(rowIndex).Cells(20).ToString()) Then
                adapterSimpan.SelectCommand.Parameters.AddWithValue("@req", "00/00/0000")
            Else
                adapterSimpan.SelectCommand.Parameters.AddWithValue("@req", DataGridView1.Rows(rowIndex).Cells(20).ToString())
            End If
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@purchase", DataGridView1.Rows(rowIndex).Cells(21).ToString())
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@po", DataGridView1.Rows(rowIndex).Cells(22).ToString())
            If String.IsNullOrEmpty(DataGridView1.Rows(rowIndex).Cells(23).ToString()) Then
                adapterSimpan.SelectCommand.Parameters.AddWithValue("@gi_ti", "00/00/0000")
            Else
                adapterSimpan.SelectCommand.Parameters.AddWithValue("@gi_ti", DataGridView1.Rows(rowIndex).Cells(23).ToString())
            End If
            If String.IsNullOrEmpty(DataGridView1.Rows(rowIndex).Cells(24).ToString()) Then
                adapterSimpan.SelectCommand.Parameters.AddWithValue("@stag", "00/00/0000")
            Else
                adapterSimpan.SelectCommand.Parameters.AddWithValue("@stag", DataGridView1.Rows(rowIndex).Cells(24).ToString())
            End If
            If String.IsNullOrEmpty(DataGridView1.Rows(rowIndex).Cells(25).ToString()) Then
                adapterSimpan.SelectCommand.Parameters.AddWithValue("@gi_da", "00/00/0000")
            Else
                adapterSimpan.SelectCommand.Parameters.AddWithValue("@gi_da", DataGridView1.Rows(rowIndex).Cells(25).ToString())
            End If
            If String.IsNullOrEmpty(DataGridView1.Rows(rowIndex).Cells(26).ToString()) Then
                adapterSimpan.SelectCommand.Parameters.AddWithValue("@mat", "00/00/0000")
            Else
                adapterSimpan.SelectCommand.Parameters.AddWithValue("@mat", DataGridView1.Rows(rowIndex).Cells(26).ToString())
            End If
            'adapterSimpan.SelectCommand.Parameters.AddWithValue("@gi_ti", Convert.ToDateTime(DataGridView1.Rows(rowIndex).Cells(23).ToString()))
            'adapterSimpan.SelectCommand.Parameters.AddWithValue("@stag", Convert.ToDateTime(DataGridView1.Rows(rowIndex).Cells(24).ToString()))
            'adapterSimpan.SelectCommand.Parameters.AddWithValue("@gi_da", Convert.ToDateTime(DataGridView1.Rows(rowIndex).Cells(25).ToString()))
            'adapterSimpan.SelectCommand.Parameters.AddWithValue("@mat", Convert.ToDateTime(DataGridView1.Rows(rowIndex).Cells(26).ToString()))
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@confirm", DataGridView1.Rows(rowIndex).Cells(27))
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@su", DataGridView1.Rows(rowIndex).Cells(28).ToString())
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@customer", DataGridView1.Rows(rowIndex).Cells(29).ToString())
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@city", DataGridView1.Rows(rowIndex).Cells(30).ToString())
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@name2", DataGridView1.Rows(rowIndex).Cells(31).ToString())
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@name1", DataGridView1.Rows(rowIndex).Cells(32).ToString())
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@item", DataGridView1.Rows(rowIndex).Cells(33))
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@so", DataGridView1.Rows(rowIndex).Cells(34))
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@cusmat", DataGridView1.Rows(rowIndex).Cells(35).ToString())
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@status", DataGridView1.Rows(rowIndex).Cells(36).ToString())
            adapterSimpan.SelectCommand.Parameters.AddWithValue("@stat", DataGridView1.Rows(rowIndex).Cells(37).ToString())

            'adapterSimpan.SelectCommand.Parameters.AddWithValue("@id", DataGridView1.Rows(rowIndex).Cells(0).Value)
            'adapterSimpan.SelectCommand.Parameters.AddWithValue("@user", DataGridView1.Rows(rowIndex).Cells(1).Value)
            'adapterSimpan.SelectCommand.Parameters.AddWithValue("@pass", DataGridView1.Rows(rowIndex).Cells(2).Value)
            adapterSimpan.SelectCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class