Imports System.Data.SqlClient
Public Class Login
    Protected Overrides Sub OnFormClosing(ByVal e As FormClosingEventArgs)
        MyBase.OnFormClosing(e)
        If Not e.Cancel AndAlso e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            MenuUtama.Show()
            Me.Hide()
        End If
    End Sub
    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar = Chr(13) Then
            TextBox2.Focus()
        End If
    End Sub
    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar = Chr(13) Then
            Button1.Focus()
        End If
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call koneksi()
        CMD = New SqlCommand("select *from TableUser where Nama_User='" & TextBox1.Text & "' and Password_User='" & TextBox2.Text & "'", CONN)
        DR = CMD.ExecuteReader
        DR.Read()
        If Not DR.HasRows Then
            MsgBox("Maaf, Login Gagal", MsgBoxStyle.Exclamation,
               "Peringatan")
            TextBox1.Clear()
            TextBox2.Clear()
            TextBox1.Focus()
        Else
            Me.Visible = False
            MenuUtama.Show()
            MenuUtama.Panel1.Text = DR.Item("Kode_User")
            MenuUtama.Panel2.Text = DR.Item("Nama_User")
            MenuUtama.Panel3.Text = DR.Item("Status_User")
            If MenuUtama.Panel3.Text = "Pegawai" Then
                MenuUtama.GolonganToolStripMenuItem.Enabled = False
                MenuUtama.PotonganToolStripMenuItem.Enabled = False
                MenuUtama.UserToolStripMenuItem.Enabled = False
                MenuUtama.JabatanToolStripMenuItem.Enabled = False
                MenuUtama.LaporanToolStripMenuItem.Enabled = False
                MenuUtama.FileToolStripMenuItem.Enabled = False
                TextBox1.Clear()
                TextBox2.Clear()
                TextBox1.Focus()
            Else
                TextBox1.Clear()
                TextBox2.Clear()
                TextBox1.Focus()
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        End
    End Sub

    Private Sub Login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class