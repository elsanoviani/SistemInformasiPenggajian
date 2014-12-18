Imports System.Data.SqlClient
Public Class GantiPassword
    Protected Overrides Sub OnFormClosing(ByVal e As FormClosingEventArgs)
        MyBase.OnFormClosing(e)
        If Not e.Cancel AndAlso e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            MenuUtama.Show()
            Me.Hide()
        End If
    End Sub
    Sub Kosongkan()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox1.Focus()
    End Sub

    Sub Form_Load()
        Call koneksi()
        TextBox1.MaxLength = 10
        TextBox2.MaxLength = 10
        TextBox3.MaxLength = 10
        TextBox2.PasswordChar = "X"
        TextBox3.PasswordChar = "X"
        TextBox4.PasswordChar = "X"
    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox4.Text <> TextBox3.Text Then
            MsgBox("password konfirmasi tidak sama")
            TextBox4.Focus()
            TextBox4.Text = ""
        Else
            If MessageBox.Show("Yakin Password akan diganti.?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Call koneksi()
                Dim edit As String = "update TableUser set Password_User='" & TextBox4.Text & "' where Kode_User='" & TextBox1.Text & "' and Password_User='" & TextBox2.Text & "'"
                CMD = New SqlCommand(edit, CONN)
                CMD.ExecuteReader()
                Call Kosongkan()
            Else
                Call Kosongkan()
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        End
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        TextBox1.MaxLength = 10
        If e.KeyChar = Chr(13) Then
            Call koneksi()
            CMD = New SqlCommand("select * from admin where kodeuser='" & TextBox1.Text & "'", CONN)
            DR.Read()
            If Not DR.Read Then
                TextBox2.Focus()
            Else
                MsgBox("Kode user salah ")
                TextBox1.Focus()
                TextBox1.Text = ""
            End If
        End If
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If e.KeyChar = Chr(13) Then
            Call koneksi()
            CMD = New SqlCommand("select * from admin where kodeuser='" & TextBox1.Text & "' and passworduser='" & TextBox2.Text & "'", CONN)
            DR.Read()
            If Not DR.Read Then
                TextBox3.Focus()
            Else
                MsgBox("password lama salah ")
                TextBox2.Focus()
                TextBox2.Text = ""
            End If
        End If
    End Sub

    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        If e.KeyChar = Chr(13) Then
            Call koneksi()
            If TextBox3.Text = "" Then
                MsgBox("Password baru belum dibuat",
                       MsgBoxStyle.Exclamation, "Peringatan")
                TextBox3.Focus()
            Else
                TextBox4.Focus()
            End If
        End If
    End Sub

    Private Sub TextBox4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        If e.KeyChar = Chr(13) Then
            Call koneksi()
            Button1.Focus()
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class