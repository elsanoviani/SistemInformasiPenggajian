Imports System.Data.SqlClient
Public Class Pegawai
    Sub Kosongkan()
        TextBox1.Clear()
        TextBox2.Clear()
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        TextBox3.Clear()
        TextBox1.Focus()
    End Sub
    Sub DataBaru()
        TextBox2.Clear()
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        TextBox3.Clear()
        TextBox2.Focus()
    End Sub
    Sub TampilGrid()
        Call Koneksi()
        DA = New SqlDataAdapter("select*from TablePegawai", CONN)
        DS = New DataSet
        DA.Fill(DS)
        DGV.DataSource = DS.Tables(0)
        DGV.ReadOnly = True
    End Sub
    Sub AmbilNama1()
        Call koneksi()
        CMD = New SqlCommand("SELECT Kode_Jabatan From TableJabatan", CONN)
        DR = CMD.ExecuteReader
        ComboBox1.Items.Clear()
        Do While DR.Read
            ComboBox1.Items.Add(DR.Item(0))
        Loop
        CMD.Dispose()
        DR.Close()
        CONN.Close()
    End Sub
    Sub AmbilNama2()
        Call koneksi()
        CMD = New SqlCommand("SELECT Golongan From TableGolongan", CONN)
        DR = CMD.ExecuteReader
        ComboBox2.Items.Clear()
        Do While DR.Read
            ComboBox2.Items.Add(DR.Item(0))
        Loop
        CMD.Dispose()
        DR.Close()
        CONN.Close()
    End Sub
    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        TextBox1.MaxLength = 25
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            CMD = New SqlCommand("select* from TablePegawai where NIP='" & TextBox1.Text & "'", CONN)
            DR = CMD.ExecuteReader
            DR.Read()
            If DR.HasRows Then
                TextBox2.Text = DR.Item("Nama_Pegawai")
                ComboBox1.Text = DR.Item("Kode_Jabatan")
                ComboBox2.Text = DR.Item("Golongan")
                ComboBox3.Text = DR.Item("Status")
                TextBox3.Text = DR.Item("Jumlah_Anak")
                TextBox2.Focus()
            Else
                Call DataBaru()
            End If
        End If
    End Sub
    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        TextBox2.MaxLength = 50
        If e.KeyChar = Chr(13) Then
            ComboBox1.Focus()
        End If
    End Sub
    Private Sub ComboBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        ComboBox1.MaxLength = 10
        If e.KeyChar = Chr(13) Then
            ComboBox2.Focus()
        End If
    End Sub
    Private Sub ComboBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        ComboBox2.MaxLength = 5
        If e.KeyChar = Chr(13) Then
            ComboBox3.Focus()
        End If
    End Sub

    Private Sub ComboBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        ComboBox3.MaxLength = 8
        If e.KeyChar = Chr(13) Then
            TextBox3.Focus()
        End If
    End Sub

    
    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        Call koneksi()
        CMD = New SqlCommand("select * from TablePegawai where Nama_Pegawai like'%" & TextBox4.Text & "%'", CONN)
        DR = CMD.ExecuteReader
        If DR.HasRows Then
            Call koneksi()
            DA = New SqlDataAdapter("select * from TablePegawai where Nama_Pegawai like'%" & TextBox4.Text & "%'", CONN)
            DS = New DataSet
            DA.Fill(DS)
            DGV.DataSource = DS.Tables(0)
        Else
            MsgBox("Nama Pegawai Tidak ditemukan", MsgBoxStyle.Critical, "INFORMASI!")
        End If
    End Sub
    Private Sub DGV_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DGV.CellMouseClick
        On Error Resume Next
        TextBox1.Text = DGV.Rows(e.RowIndex).Cells(0).Value
        TextBox2.Text = DGV.Rows(e.RowIndex).Cells(1).Value
        ComboBox1.Text = DGV.Rows(e.RowIndex).Cells(2).Value
        ComboBox2.Text = DGV.Rows(e.RowIndex).Cells(3).Value
        ComboBox3.Text = DGV.Rows(e.RowIndex).Cells(4).Value
        TextBox3.Text = DGV.Rows(e.RowIndex).Cells(5).Value
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Close()
    End Sub

    Private Sub Pegawai_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call koneksi()
        Call TampilGrid()
        Call AmbilNama1()
        Call AmbilNama2()
    End Sub

    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Call koneksi()
        Dim edit As String = "update TablePegawai set Nama_Pegawai='" & TextBox2.Text & "',Kode_Jabatan='" & ComboBox1.Text & "',Golongan='" & ComboBox2.Text & "',Status='" & ComboBox3.Text & "',Jumlah_Anak='" & TextBox3.Text & "' where NIP='" & TextBox1.Text & "'"
        CMD = New SqlCommand(edit, CONN)
        CMD.ExecuteNonQuery()
        Call Kosongkan()
        Call TampilGrid()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call Kosongkan()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MsgBox("NIP Harus Diisi", MsgBoxStyle.Critical, "WARNING!")
            Exit Sub
        Else
            If MessageBox.Show("Hapus Data ini...?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Call koneksi()
                Dim hapus As String = "delete from TablePegawai where NIP='" & TextBox1.Text & "'"
                CMD = New SqlCommand(hapus, CONN)
                CMD.ExecuteReader()
                Call Kosongkan()
                Call TampilGrid()
            Else
                Call Kosongkan()
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or ComboBox1.Text = "" Or ComboBox2.Text = "" Or ComboBox3.Text = "" Or TextBox3.Text = "" Then
            MsgBox("Data Belum Lengkap", MsgBoxStyle.Critical, "WARNING!")
            Exit Sub
        Else
            Call koneksi()
            CMD = New SqlCommand("select * from TablePegawai where NIP ='" & TextBox1.Text & "'", CONN)
            DR = CMD.ExecuteReader
            DR.Read()
            If Not DR.HasRows Then
                Call koneksi()
                Dim simpan As String = "insert into TablePegawai values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & ComboBox1.Text & "','" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & TextBox3.Text & "')"
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
End Class