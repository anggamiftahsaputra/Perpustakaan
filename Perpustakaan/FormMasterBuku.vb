Imports System.Data.SqlClient
Public Class FormMasterBuku
    Sub KondisiAwal()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        ComboBox1.Text = ""
        TextBox6.Text = ""
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        ComboBox1.Enabled = False
        TextBox6.Enabled = False
        Button1.Text = "INPUT"
        Button2.Text = "EDIT"
        Button3.Text = "HAPUS"
        Button4.Text = "TUTUP"
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        ComboBox1.MaxLength = 4

        'ComboBox1.Items.Add("2018")
        'ComboBox1.Items.Add("2017")
        'For i = DateTime.Now.Year - 5 To DateTime.Now.Year
        '    ComboBox1.Items.Add(i)
        'Next
        Dim Tahun As String
        Tahun = Date.Now().Year
        ComboBox1.Items.Clear()
        With ComboBox1
            While Tahun >= 2014
                .Items.Add(Tahun)
                Tahun = Tahun - 1

            End While
        End With

        Call koneksi()
        Da = New SqlDataAdapter("Select KodeBuku, JudulBuku, PengarangBuku, PenerbitBuku, TahunBuku, StokBuku from TBL_BUKU", Conn)
        Ds = New DataSet
        Da.Fill(Ds, "TBL_BUKU")
        DataGridView1.DataSource = (Ds.Tables("TBL_BUKU"))
    End Sub
    Sub SiapIsi()
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        ComboBox1.Enabled = True
        TextBox6.Enabled = True
    End Sub
    Private Sub FormMasterBuku_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call KondisiAwal()
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Button1.Text = "INPUT" Then
            Button1.Text = "SIMPAN"
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Text = "BATAL"
            Call SiapIsi()
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or ComboBox1.Text = "" Or TextBox6.Text = "" Then
                MsgBox("Pastikan Isi Semua Field")
            Else
                Call koneksi()
                Dim SimpanData As String = "insert into tbl_buku values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & ComboBox1.Text & "','" & TextBox6.Text & "')"
                Cmd = New SqlCommand(SimpanData, Conn)
                Cmd.ExecuteNonQuery()
                MsgBox("Data Berhasil Di Simpan")
                Call KondisiAwal()
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Button2.Text = "EDIT" Then
            Button1.Enabled = False
            Button2.Text = "SIMPAN"
            Button3.Enabled = False
            Button4.Text = "BATAL"
            Call SiapIsi()
        Else

            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or ComboBox1.Text = "" Or TextBox6.Text = "" Then
                MsgBox("Pastikan Isi Semua Field")
            Else
                Call koneksi()
                Dim EditData As String = "Update tbl_buku set JudulBuku='" & TextBox2.Text & "',PengarangBuku ='" & TextBox3.Text & "',PenerbitBuku='" & TextBox4.Text & "',TahunBuku='" & ComboBox1.Text & "',StokBuku='" & TextBox6.Text & "' where KodeBuku='" & TextBox1.Text & "'"
                Cmd = New SqlCommand(EditData, Conn)
                Cmd.ExecuteNonQuery()
                MsgBox("Data Berhasil Di Edit")
                Call KondisiAwal()
            End If
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call SiapIsi()
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or ComboBox1.Text = "" Or TextBox6.Text = "" Then
            MsgBox("Pastikan Isi Semua Field")
        Else
            Call koneksi()
            Dim HapusData As String = "delete tbl_buku where KodeBuku='" & TextBox1.Text & "'"
            Cmd = New SqlCommand(HapusData, Conn)
            Cmd.ExecuteNonQuery()
            MsgBox("Data Berhasil Di Hapus")
            Call KondisiAwal()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Button4.Text = "TUTUP" Then
            Me.Close()
        Else
            Call KondisiAwal()
        End If
    End Sub
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call koneksi()
            Cmd = New SqlCommand("Select * From tbl_buku where KodeBuku='" & TextBox1.Text & "'", Conn)
            Rd = Cmd.ExecuteReader
            Rd.Read()
            If Rd.HasRows Then
                TextBox2.Text = Rd.Item("JudulBuku")
                TextBox3.Text = Rd.Item("PengarangBuku")
                TextBox4.Text = Rd.Item("PenerbitBuku")
                ComboBox1.Text = Rd.Item("TahunBuku")
                TextBox6.Text = Rd.Item("StokBuku")
            Else
                MsgBox("Data Tidak Ada!")
            End If
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub
End Class