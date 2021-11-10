Imports System.Data.SqlClient

Public Class Input_BoX
    Private Sub Input_BoX_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TopMost = True
        txt_PlanCode_DateCode.Text = ""
        txt_QRCode.Text = ""
        Txt_Barcode.Text = ""
        btn_Cek.Visible = False
    End Sub

    Private Sub Txt_QRCode_TextChanged(sender As Object, e As PreviewKeyDownEventArgs) Handles txt_QRCode.PreviewKeyDown ', txt_QRCode.TextChanged
        'If txt_QRCode.TextLength >= 470 And (e.KeyData = Keys.Enter Or e.KeyData = Keys.Tab) Then
        If e.KeyData = Keys.Enter Or e.KeyData = Keys.Tab Then

            If txt_QRCode.Text.Length >= 50 Then
                MsgBox("Wrong QRCode !")
                Exit Sub
            End If

            If txt_QRCode.Text.Length = 13 Then
                MsgBox("Pls QRCode instead of EAN13 Barcode!")
                Exit Sub
            End If

            Combo_PlanCode.Text = ""
            Txt_Barcode.Text = ""
            'txt_Ref.Text = txt_QRCode.Text.Substring(23, 8)
            txt_PlanCode_DateCode.Text = txt_QRCode.Text.Substring(7, 7)
            txt_PlanCode_DateCode.Select()
            SendKeys.Send("{ENTER}")
            'txt_QRCode.SelectAll()
        End If

    End Sub

    Private Sub Txt_Barcode_TextChanged(sender As Object, e As PreviewKeyDownEventArgs) Handles Txt_Barcode.PreviewKeyDown
        If e.KeyData = Keys.Enter Or e.KeyData = Keys.Tab Then
            If Txt_Barcode.TextLength >= 10 Then
                txt_QRCode.Text = ""
                'txt_Ref.Text = Txt_Barcode.Text.Substring(6, 5)
                If Combo_PlanCode.Text = "" Then
                    MsgBox("Please Select PlanCode !")
                Else
                    txt_PlanCode_DateCode.Text = Combo_PlanCode.Text & Txt_Barcode.Text.Substring(1, 5)
                    txt_PlanCode_DateCode.Select()
                    SendKeys.Send("{ENTER}")
                    'Txt_Barcode.SelectAll()
                End If

            End If
        End If

    End Sub
    Private Sub Txt_PlanCode_DateCode_TextChanged(sender As Object, e As PreviewKeyDownEventArgs) Handles txt_PlanCode_DateCode.PreviewKeyDown
        If Combo_PlanCode.Text <> "" Then
            btn_Cek.Visible = True
        End If
        If (e.KeyData = Keys.Enter Or e.KeyData = Keys.Tab) Then
            If txt_PlanCode_DateCode.TextLength >= 7 Then
                Btn_Cek_Click(sender, e)
            End If
        End If
    End Sub

    Private Sub Btn_Cek_Click(sender As Object, e As EventArgs) Handles btn_Cek.Click

        If txt_PlanCode_DateCode.TextLength >= 7 Then
            'txt_PlanCode_DateCode.Text = Microsoft.VisualBasic.Left(txt_PlanCode_DateCode.Text, 7)
            Me.Close()
        Else
            MsgBox("Plan Code and Date Code incorrect !")
        End If
    End Sub

    Private Sub Combo_PlanCode_Click(sender As Object, e As EventArgs) Handles Combo_PlanCode.Click
        Dim Adap = New SqlDataAdapter("Select [Plant Code] from MasterPlantCode", Main.koneksi)
        Dim dt = New DataTable
        Try
            Adap.Fill(dt)
            Combo_PlanCode.DataSource = dt
            Combo_PlanCode.ValueMember = "Plant Code"
            Combo_PlanCode.DisplayMember = "Plant Code"
        Catch ex As Exception

        End Try

    End Sub

End Class