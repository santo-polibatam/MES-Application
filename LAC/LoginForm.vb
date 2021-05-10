<<<<<<< HEAD
﻿Imports System.Data.SqlClient

Public Class LoginForm

    Public strHostName As String = System.Net.Dns.GetHostName()

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call Main.koneksi_db()

        Try
            Dim sc As New SqlCommand("select * from login where username=@username and password=@password", Main.koneksi)
            sc.Parameters.Add("@username", SqlDbType.VarChar).Value = textboxusername.Text
            sc.Parameters.Add("@password", SqlDbType.VarChar).Value = textboxpassword.Text
            Dim adapter As New SqlDataAdapter(sc)
            Dim dt As New DataTable()

            adapter.Fill(dt)

            If dt.Rows.Count() <= 0 Then
                MsgBox("Invalid Username Or Password")
            Else
                Main.role = textboxusername.Text
                Me.Hide()
                Main.Show()

            End If

            If Main.role = "admin" Then
                Main.AccessAdmin()
            Else
                Main.AccessOperator()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub textboxpassword_KeyPress(sender As Object, e As KeyPressEventArgs) Handles textboxpassword.KeyPress
        If e.KeyChar = Chr(13) Then
            Call Main.koneksi_db()

            Try
                Dim sc As New SqlCommand("select * from login where username=@username and password=@password", Main.koneksi)
                sc.Parameters.Add("@username", SqlDbType.VarChar).Value = textboxusername.Text
                sc.Parameters.Add("@password", SqlDbType.VarChar).Value = textboxpassword.Text
                Dim adapter As New SqlDataAdapter(sc)
                Dim dt As New DataTable()

                adapter.Fill(dt)

                If dt.Rows.Count() <= 0 Then
                    MsgBox("Invalid Username Or Password")
                Else
                    Main.role = textboxusername.Text
                    Main.Show()
                    Me.Hide()
                End If

                If Main.role = "admin" Then
                    Main.AccessAdmin()
                Else
                    Main.AccessOperator()
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub

    Private Sub LoginForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Form2.Hide()
        Me.CenterToScreen()
        Me.Label3.Text = "V " & Application.ProductVersion
        Me.Text = "LoginForm - " & strHostName
    End Sub

    Private Sub Textboxusername_KeyPress(sender As Object, e As KeyPressEventArgs) Handles textboxusername.KeyPress
        If e.KeyChar = Chr(13) Then
            Call Main.koneksi_db()
            If textboxusername.Text = "operator" Then
                Main.role = textboxusername.Text
                Main.Show()
                Me.Hide()
            Else
                MsgBox("Password Cannot be null")
            End If
        End If
    End Sub

    Private Sub Textboxusername_TextChanged(sender As Object, e As EventArgs) Handles textboxusername.TextChanged
        If textboxusername.Text = "operator" Then
            textboxpassword.Enabled = False
            Button1.Enabled = False
        Else
            textboxpassword.Enabled = True
            Button1.Enabled = True
        End If

        If Main.role = "admin" Then
            Main.AccessAdmin()
        Else
            Main.AccessOperator()
        End If
    End Sub

    Private Sub Textboxpassword_TextChanged(sender As Object, e As EventArgs) Handles textboxpassword.TextChanged

    End Sub
=======
﻿Imports System.Data.SqlClient

Public Class LoginForm

    Public strHostName As String = System.Net.Dns.GetHostName()

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call Main.koneksi_db()

        Try
            Dim sc As New SqlCommand("select * from login where username=@username and password=@password", Main.koneksi)
            sc.Parameters.Add("@username", SqlDbType.VarChar).Value = textboxusername.Text
            sc.Parameters.Add("@password", SqlDbType.VarChar).Value = textboxpassword.Text
            Dim adapter As New SqlDataAdapter(sc)
            Dim dt As New DataTable()

            adapter.Fill(dt)

            If dt.Rows.Count() <= 0 Then
                MsgBox("Invalid Username Or Password")
            Else
                Main.role = textboxusername.Text
                Me.Hide()
                Main.Show()

            End If

            If Main.role = "admin" Then
                Main.AccessAdmin()
            Else
                Main.AccessOperator()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub textboxpassword_KeyPress(sender As Object, e As KeyPressEventArgs) Handles textboxpassword.KeyPress
        If e.KeyChar = Chr(13) Then
            Call Main.koneksi_db()

            Try
                Dim sc As New SqlCommand("select * from login where username=@username and password=@password", Main.koneksi)
                sc.Parameters.Add("@username", SqlDbType.VarChar).Value = textboxusername.Text
                sc.Parameters.Add("@password", SqlDbType.VarChar).Value = textboxpassword.Text
                Dim adapter As New SqlDataAdapter(sc)
                Dim dt As New DataTable()

                adapter.Fill(dt)

                If dt.Rows.Count() <= 0 Then
                    MsgBox("Invalid Username Or Password")
                Else
                    Main.role = textboxusername.Text
                    Main.Show()
                    Me.Hide()
                End If

                If Main.role = "admin" Then
                    Main.AccessAdmin()
                Else
                    Main.AccessOperator()
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub

    Private Sub LoginForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Form2.Hide()
        Me.CenterToScreen()
        Me.Label3.Text = "V " & Application.ProductVersion
        Me.Text = "LoginForm - " & strHostName
    End Sub

    Private Sub Textboxusername_KeyPress(sender As Object, e As KeyPressEventArgs) Handles textboxusername.KeyPress
        If e.KeyChar = Chr(13) Then
            Call Main.koneksi_db()
            If textboxusername.Text = "operator" Then
                Main.role = textboxusername.Text
                Main.Show()
                Me.Hide()
            Else
                MsgBox("Password Cannot be null")
            End If
        End If
    End Sub

    Private Sub Textboxusername_TextChanged(sender As Object, e As EventArgs) Handles textboxusername.TextChanged
        If textboxusername.Text = "operator" Then
            textboxpassword.Enabled = False
            Button1.Enabled = False
        Else
            textboxpassword.Enabled = True
            Button1.Enabled = True
        End If

        If Main.role = "admin" Then
            Main.AccessAdmin()
        Else
            Main.AccessOperator()
        End If
    End Sub

    Private Sub Textboxpassword_TextChanged(sender As Object, e As EventArgs) Handles textboxpassword.TextChanged

    End Sub
>>>>>>> 5d2bc5751623f3413e8c53f892b2d6831328829d
End Class