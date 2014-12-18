Imports System.Data.SqlClient
Public Class User
    Protected Overrides Sub OnFormClosing(ByVal e As FormClosingEventArgs)
        MyBase.OnFormClosing(e)
        If Not e.Cancel AndAlso e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            MenuUtama.Show()
            Me.Hide()
        End If
    End Sub
    Sub Kosongkan()
        TextBox2.Clear()
        TextBox3.Clear()
        ComboBox1.Text = ""
        TextBox2.Focus()
    End Sub
    Sub DataBaru()
        TextBox2.Clear()
        TextBox3.Clear()
        ComboBox1.Text = ""
        TextBox2.Focus()
    End Sub
    Sub TampilGrid()
        Call koneksi()
        DA = New SqlDataAdapter("select*from TableUser", CONN)
        DS = New DataSet
        DA.Fill(DS)
        DGV.DataSource = DS.Tables(0)
        DGV.ReadOnly = True
    End Sub
    Sub Otomatis()
        CMD = New SqlCommand("select * from TableUser order by Kode_User desc", CONN)
        DR = CMD.ExecuteReader
        DR.Read()

        If Not DR.HasRows Then
            Label5.Text = "KU001"
        Else
            Label5.Text = Val(Microsoft.VisualBasic.Mid(DR.Item("Kode_User").ToString, 3, 3)) + 1

            If Len(Label5.Text) = 1 Then
                Label5.Text = "KU00" & Label5.Text & ""
            ElseIf Len(Label5.Text) = 2 Then
                Label5.Text = "KU0" & Label5.Text & ""
            ElseIf Len(Label5.Text) = 3 Then
                Label5.Text = "KU" & Label5.Text & ""
            End If
        End If
    End Sub

    Private Sub label5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar = Chr(13) Then
            Call koneksi()
            CMD = New SqlCommand("select* from TableUser where Kode_User='" & Label5.Text & "'", CONN)
            DR = CMD.ExecuteReader
            DR.Read()
            If DR.HasRows Then
                TextBox2.Text = DR.Item("Nama_User")
                TextBox3.Text = DR.Item("Password")
                ComboBox1.Text = DR.Item("Status_User")
                TextBox2.Focus()
            Else
                Call DataBaru()
            End If
        End If
    End Sub
    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        TextBox2.MaxLength = 30
        If e.KeyChar = Chr(13) Then
            TextBox3.Focus()
        End If
    End Sub
    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        TextBox3.MaxLength = 10
        If e.KeyChar = Chr(13) Then
            ComboBox1.Focus()
        End If
    End Sub
    Private Sub _combobox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox1.KeyPress
        ComboBox1.MaxLength = 15
        If e.KeyChar = Chr(13) Then
            Button1.Focus()
        End If
    End Sub

    Private Sub Jabatan_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Call Otomatis()
    End Sub

    Private Sub Jabatan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call TampilGrid()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Label5.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("Data Belum Lengkap", MsgBoxStyle.Critical, "WARNING!")
            Exit Sub
        Else
            Call koneksi()
            CMD = New SqlCommand("select * from TableUser where Kode_User ='" & Label5.Text & "'", CONN)
            DR = CMD.ExecuteReader
            DR.Read()
            If Not DR.HasRows Then
                Call koneksi()
                Dim simpan As String = "insert into TableUser values('" & Label5.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & ComboBox1.Text & "')"
                CMD = New SqlCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
            Else
                MsgBox("Maaf, Data dengan kode tersebut telah ada",
                        MsgBoxStyle.Exclamation, "Peringatan")
                Call Kosongkan()
            End If
            Call Otomatis()
            Call Kosongkan()
            Call TampilGrid()
        End If
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Label5.Text = "" Then
            MsgBox("Kode Harus Diisi", MsgBoxStyle.Critical, "WARNING!")
            Exit Sub
        Else
            If MessageBox.Show("Hapus Data ini...?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Call koneksi()
                Dim hapus As String = "delete from TableUser where Kode_User='" & Label5.Text & "'"
                CMD = New SqlCommand(hapus, CONN)
                CMD.ExecuteReader()
                Call Kosongkan()
                Call TampilGrid()
                Call Otomatis()
            Else
                Call Kosongkan()
                Call Otomatis()
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call Kosongkan()
        Call koneksi()
        Call Otomatis()
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        Call koneksi()
        CMD = New SqlCommand("select*from TableUser where Nama_User like'%" & TextBox5.Text & "%'", CONN)
        DR = CMD.ExecuteReader
        If DR.HasRows Then
            Call koneksi()
            DA = New SqlDataAdapter("select*from TableUser where Nama_USer like'%" & TextBox5.Text & "%'", CONN)
            DS = New DataSet
            DA.Fill(DS)
            DGV.DataSource = DS.Tables(0)
        Else
            MsgBox("Nama Jabatan Tidak ditemukan", MsgBoxStyle.Critical, "INFORMASI!")
        End If
    End Sub
    Private Sub DGV_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DGV.CellMouseClick
        Label5.Text = DGV.Rows(e.RowIndex).Cells(0).Value
        TextBox2.Text = DGV.Rows(e.RowIndex).Cells(1).Value
        TextBox3.Text = DGV.Rows(e.RowIndex).Cells(2).Value
        ComboBox1.Text = DGV.Rows(e.RowIndex).Cells(3).Value
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call koneksi()
        Dim edit As String = "update TableUser set Nama_User='" & TextBox2.Text & "',Password_User='" & TextBox3.Text & "',Status_User='" & ComboBox1.Text & "' where Kode_User='" & Label5.Text & "'"
        CMD = New SqlCommand(edit, CONN)
        CMD.ExecuteNonQuery()
        Call Kosongkan()
        Call TampilGrid()
        Call Otomatis()
    End Sub

    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
        MenuUtama.Show()
    End Sub
End Class


