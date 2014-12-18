Imports System.Data.SqlClient
Public Class Golongan
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
        TextBox5.Clear()
        TextBox1.Focus()
    End Sub
    Sub DataBaru()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox1.Focus()
    End Sub
    Sub TampilGrid()
        Call koneksi()
        DA = New SqlDataAdapter("select*from TableGolongan", CONN)
        DS = New DataSet
        DA.Fill(DS)
        DGV.DataSource = DS.Tables(0)
        DGV.ReadOnly = True
    End Sub

    Private Sub textbox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar = Chr(13) Then
            Call koneksi()
            CMD = New SqlCommand("select* from TableGolongan where Golongan='" & TextBox1.Text & "'", CONN)
            DR = CMD.ExecuteReader
            DR.Read()
            If DR.HasRows Then
                TextBox2.Text = DR.Item("Tj_Suami_Istri")
                TextBox3.Text = DR.Item("Tj_Anak")
                TextBox4.Text = DR.Item("Tj_Beras")
                TextBox5.Text = DR.Item("Askes")
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
            TextBox4.Focus()
        End If
    End Sub
    Private Sub _TextBox4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        TextBox4.MaxLength = 15
        If e.KeyChar = Chr(13) Then
            TextBox5.Focus()
        End If
    End Sub

    Private Sub Jabatan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call koneksi()
        Call TampilGrid()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            MsgBox("Data Belum Lengkap", MsgBoxStyle.Critical, "WARNING!")
            Exit Sub
        Else
            Call koneksi()
            CMD = New SqlCommand("select * from TableGolongan where Golongan ='" & TextBox1.Text & "'", CONN)
            DR = CMD.ExecuteReader
            DR.Read()
            If Not DR.HasRows Then
                Call koneksi()
                Dim simpan As String = "insert into TableGolongan values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "')"
                CMD = New SqlCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
            Else
                MsgBox("Maaf, Data dengan kode tersebut telah ada",
                        MsgBoxStyle.Exclamation, "Peringatan")
                Call Kosongkan()
            End If
            Call Kosongkan()
            Call TampilGrid()
        End If
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MsgBox("Kode Harus Diisi", MsgBoxStyle.Critical, "WARNING!")
            Exit Sub
        Else
            If MessageBox.Show("Hapus Data ini...?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Call koneksi()
                Dim hapus As String = "delete from TableGolongan where Golongan='" & TextBox1.Text & "'"
                CMD = New SqlCommand(hapus, CONN)
                CMD.ExecuteReader()
                Call Kosongkan()
                Call TampilGrid()
            Else
                Call Kosongkan()
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call Kosongkan()
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged
        Call koneksi()
        CMD = New SqlCommand("select*from TableGolongan where Golongan like'%" & TextBox6.Text & "%'", CONN)
        DR = CMD.ExecuteReader
        If DR.HasRows Then
            Call koneksi()
            DA = New SqlDataAdapter("select*from TableGolongan where Golongan like'%" & TextBox6.Text & "%'", CONN)
            DS = New DataSet
            DA.Fill(DS)
            DGV.DataSource = DS.Tables(0)
        Else
            MsgBox("Nama Jabatan Tidak ditemukan", MsgBoxStyle.Critical, "INFORMASI!")
        End If
    End Sub
    Private Sub DGV_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DGV.CellMouseClick
        On Error Resume Next
        TextBox1.Text = DGV.Rows(e.RowIndex).Cells(0).Value
        TextBox2.Text = DGV.Rows(e.RowIndex).Cells(1).Value
        TextBox3.Text = DGV.Rows(e.RowIndex).Cells(2).Value
        TextBox4.Text = DGV.Rows(e.RowIndex).Cells(3).Value
        TextBox5.Text = DGV.Rows(e.RowIndex).Cells(4).Value
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call koneksi()
        Dim edit As String = "update TableGolongan set Nama_Golongan='" & TextBox2.Text & "',Tj_Suami_Istri='" & TextBox3.Text & "',Tj_Anak='" & TextBox4.Text & "',Tj_Beras='" & TextBox4.Text & "',Akses='" & TextBox5.Text & "' where Golongan='" & TextBox1.Text & "'"
        CMD = New SqlCommand(edit, CONN)
        CMD.ExecuteNonQuery()
        Call Kosongkan()
        Call TampilGrid()
    End Sub

    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
        MenuUtama.Show()
    End Sub

    Private Sub TextBox5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress
        TextBox5.MaxLength = 10
        If e.KeyChar = Chr(13) Then
            Button1.Focus()
        End If
    End Sub

End Class