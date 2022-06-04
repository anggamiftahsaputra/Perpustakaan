Public Class MenuUtama
    Private Sub KeluarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles KeluarToolStripMenuItem.Click
        End
    End Sub
    Sub Terkunci()
        LoginToolStripMenuItem.Enabled = True
        LogoutToolStripMenuItem.Enabled = False
        MasterToolStripMenuItem.Enabled = False
        TransaksiToolStripMenuItem.Enabled = False
        LaporanToolStripMenuItem.Enabled = False
        UtilityToolStripMenuItem.Enabled = False
    End Sub

    Private Sub LoginToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoginToolStripMenuItem.Click
        FormLogin.ShowDialog()

    End Sub

    Private Sub MenuUtama_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call Terkunci()
        STLabel10.Text = Today
    End Sub

    Private Sub FIleToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FIleToolStripMenuItem.Click

    End Sub

    Private Sub LogoutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogoutToolStripMenuItem.Click
        Call Terkunci()
    End Sub

    Private Sub PetugasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PetugasToolStripMenuItem.Click
        FormMasterPetugas.ShowDialog()
    End Sub

    Private Sub AnggotaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AnggotaToolStripMenuItem.Click
        FormMasterAnggota.ShowDialog()
    End Sub

    Private Sub BukuToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BukuToolStripMenuItem.Click
        FormMasterBuku.ShowDialog()
    End Sub

    Private Sub PeminjamanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PeminjamanToolStripMenuItem.Click
        FormTransPinjam.ShowDialog()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        STLabel8.Text = TimeOfDay
    End Sub
End Class
