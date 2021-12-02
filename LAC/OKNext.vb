Public Class OKNext
    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form1.OKNEXT_Var = 1
        DataGridUpdateDb.OKNEXT_Var = 1
        Me.Close()
    End Sub

    Public Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form1.OKNEXT_Var = 0
        DataGridUpdateDb.OKNEXT_Var = 0
        Me.Close()

    End Sub
End Class