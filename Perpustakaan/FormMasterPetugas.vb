Imports System.Data.SqlClient
Public Class FormMasterPetugas
    Sub KondisiAwal()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        ComboBox1.Text = ""
        Call koneksi()
        Da = New SqlDataAdapter("Select KodePetugas, NamaPetugas, LevelPetugas from TBL_PETUGAS", Conn)
        Ds = New DataSet
        Da.Fill(Ds, "tbl_petugas")
        DataGridView1.DataSource = (Ds.Tables("tbl_petugas"))
        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("ADMIN")
        ComboBox1.Items.Add("USER")
        TextBox3.PasswordChar = "*"
    End Sub
    Sub SiapIsi()
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        ComboBox1.Enabled = True
    End Sub


    Private Sub FormMasterPetugas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call KondisiAwal()
        TextBox1.MaxLength = 6
        TextBox2.MaxLength = 50
        TextBox3.MaxLength = 30
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("Pastikan Isi Semua Field")
        Else
            Call koneksi()
            Dim SimpanData As String = "insert into tbl_petugas values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & ComboBox1.Text & "')"
            Cmd = New SqlCommand(SimpanData, Conn)
            Cmd.ExecuteNonQuery()
            MsgBox("Data Berhasil Di Simpan")
            Call KondisiAwal()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("Pastikan Isi Semua Field")
        Else
            Call koneksi()
            Dim EditData As String = "Update tbl_petugas set NamaPetugas='" & TextBox2.Text & "',passwordpetugas='" & TextBox3.Text & "',levelpetugas='" & ComboBox1.Text & "' where kodepetugas='" & TextBox1.Text & "'"
            Cmd = New SqlCommand(EditData, Conn)
            Cmd.ExecuteNonQuery()
            MsgBox("Data Berhasil Di Edit")
            Call KondisiAwal()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ''If Button3.Text = "HAPUS" Then
        '    Button1.Enabled = False
        '    Button2.Enabled = False
        '    Button3.Text = "SIMPAN"
        '    Button4.Text = "BATAL"
        '    Call SiapIsi()
        'Else
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("Pastikan Isi Semua Field")
        Else
            Call koneksi()
            Dim HapusData As String = "delete tbl_petugas where kodepetugas='" & TextBox1.Text & "'"
            Cmd = New SqlCommand(HapusData, Conn)
            Cmd.ExecuteNonQuery()
            MsgBox("Data Berhasil Di Hapus")
            Call KondisiAwal()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call koneksi()
            Cmd = New SqlCommand("Select * From tbl_Petugas where kodepetugas='" & TextBox1.Text & "'", Conn)
            Rd = Cmd.ExecuteReader
            Rd.Read()
            If Rd.HasRows Then
                TextBox2.Text = Rd.Item("NamaPetugas")
                TextBox3.Text = Rd.Item("PasswordPetugas")
                ComboBox1.Text = Rd.Item("LevelPetugas")
            Else
                MsgBox("Data Tidak Ada!")
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class