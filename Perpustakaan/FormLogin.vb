Imports System.Data.SqlClient
Public Class FormLogin
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call koneksi()
        Cmd = New SqlCommand("Select * from TBL_PETUGAS where KodePetugas='" & TextBox1.Text & "' and PasswordPetugas='" & TextBox2.Text & "'", Conn)
        Rd = Cmd.ExecuteReader
        Rd.Read()
        If Rd.HasRows Then
            Me.Close()
            Call BukaKunci()
            MenuUtama.STLabel2.Text = Rd!KodePetugas
            MenuUtama.STLabel4.Text = Rd!NamaPetugas
            MenuUtama.STLabel6.Text = Rd!LevelPetugas


        Else
            MsgBox("Kode Petugas Atau Password Salah!!")
        End If
    End Sub
    Sub BukaKunci()
        MenuUtama.LoginToolStripMenuItem.Enabled = False
        MenuUtama.LogoutToolStripMenuItem.Enabled = True
        MenuUtama.MasterToolStripMenuItem.Enabled = True
        MenuUtama.TransaksiToolStripMenuItem.Enabled = True
        MenuUtama.LaporanToolStripMenuItem.Enabled = True
        MenuUtama.UtilityToolStripMenuItem.Enabled = True
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub FormLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = "PTG001"
        TextBox2.Text = "ADMIN"
        TextBox2.PasswordChar = "*"
    End Sub
End Class