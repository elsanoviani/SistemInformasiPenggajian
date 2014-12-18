Imports System.Data.SqlClient
Public Class MenuUtama
    Protected Overrides Sub OnFormClosing(ByVal e As FormClosingEventArgs)
        MyBase.OnFormClosing(e)
        If Not e.Cancel AndAlso e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            If MessageBox.Show("Apakah Anda yakin ingin Keluar dari Program", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                End
            End If
        End If
    End Sub
    Private Sub PetugasToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UserToolStripMenuItem.Click
        User.Show()
    End Sub

    Private Sub UserToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GolonganToolStripMenuItem.Click
        Golongan.Show()
    End Sub

    Private Sub PegawaiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles JabatanToolStripMenuItem.Click
        Jabatan.Show()
    End Sub

    Private Sub GajiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PotonganToolStripMenuItem.Click

        Potongan.Show()
    End Sub

    Private Sub PegawaiToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PegawaiToolStripMenuItem1.Click

        Pegawai.Show()
    End Sub

    Private Sub KeluarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KeluarToolStripMenuItem.Click
        If MessageBox.Show("Anda Yakin Ingin Keluar dari Program Ini?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            Me.Hide()
            Login.Show()
            Me.Close()
        End If
    End Sub

    Private Sub PenggajianToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PenggajianToolStripMenuItem.Click

        Penggajian.Show()

    End Sub

    Private Sub GantiPasswordToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GantiPasswordToolStripMenuItem.Click
        GantiPassword.Show()
    End Sub

    Private Sub DataPegawaiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataPegawaiToolStripMenuItem.Click
        LPegawai.Show()
    End Sub

    Private Sub GroupBox2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    

    Private Sub ToolStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs)

    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub DataGajiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGajiToolStripMenuItem.Click
        Laporan.Show()
    End Sub

    Private Sub FileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FileToolStripMenuItem.Click

    End Sub
End Class