Imports System.Data.SqlClient
Public Class Jabatan
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
        TextBox4.Text = ""
        TextBox2.Focus()
    End Sub
    Sub DataBaru()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Text = ""
        TextBox2.Focus()
    End Sub
    Sub TampilGrid()
        Call Koneksi()
        DA = New SqlDataAdapter("select*from TableJabatan", CONN)
        DS = New DataSet
        DA.Fill(DS)
        DGV.DataSource = DS.Tables(0)
        DGV.ReadOnly = True
    End Sub
    Sub Otomatis()
        CMD = New SqlCommand("select * from TableJabatan order by Kode_Jabatan desc", CONN)
        DR = CMD.ExecuteReader
        DR.Read()

        If Not DR.HasRows Then
            Label5.Text = "KJ001"
        Else
            Label5.Text = Val(Microsoft.VisualBasic.Mid(DR.Item("Kode_Jabatan").ToString, 3, 3)) + 1

            If Len(Label5.Text) = 1 Then
                Label5.Text = "KJ00" & Label5.Text & ""
            ElseIf Len(Label5.Text) = 2 Then
                Label5.Text = "KJ0" & Label5.Text & ""
            ElseIf Len(Label5.Text) = 3 Then
                Label5.Text = "KJ" & Label5.Text & ""
            End If
        End If
    End Sub

    Private Sub label5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            CMD = New SqlCommand("select* from TableJabatan where Kode_Jabatan='" & Label5.Text & "'", CONN)
            DR = CMD.ExecuteReader
            DR.Read()
            If DR.HasRows Then
                TextBox2.Text = DR.Item("Nama_Jabatan")
                TextBox3.Text = DR.Item("Gaji_Pokok")
                TextBox4.Text = DR.Item("Tj_Jabatan")
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
            Button1.Focus()
        End If
    End Sub

    Private Sub Jabatan_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Call Otomatis()
    End Sub

    Private Sub Jabatan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Koneksi()
        Call TampilGrid()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Label5.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            MsgBox("Data Belum Lengkap", MsgBoxStyle.Critical, "WARNING!")
            Exit Sub
        Else
            Call Koneksi()
            CMD = New SqlCommand("select * from TableJabatan where Kode_Jabatan ='" & Label5.Text & "'", CONN)
            DR = CMD.ExecuteReader
            DR.Read()
            If Not DR.HasRows Then
                Call Koneksi()
                Dim simpan As String = "insert into TableJabatan values('" & Label5.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "')"
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
        If label5.Text = "" Then
            MsgBox("Kode Harus Diisi", MsgBoxStyle.Critical, "WARNING!")
            Exit Sub
        Else
            If MessageBox.Show("Hapus Data ini...?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Call Koneksi()
                Dim hapus As String = "delete from TableJabatan where Kode_Jabatan='" & Label5.Text & "'"
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

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        Call Koneksi()
        CMD = New SqlCommand("select*from TableJabatan where Nama_Jabatan like'%" & TextBox5.Text & "%'", CONN)
        DR = CMD.ExecuteReader
        If DR.HasRows Then
            Call Koneksi()
            DA = New SqlDataAdapter("select*from TableJabatan where Nama_Jabatan like'%" & TextBox5.Text & "%'", CONN)
            DS = New DataSet
            DA.Fill(DS)
            DGV.DataSource = DS.Tables(0)
        Else
            MsgBox("Nama Jabatan Tidak ditemukan", MsgBoxStyle.Critical, "INFORMASI!")
        End If
    End Sub
    Private Sub DGV_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DGV.CellMouseClick
        On Error Resume Next
        Label5.Text = DGV.Rows(e.RowIndex).Cells(0).Value
        TextBox2.Text = DGV.Rows(e.RowIndex).Cells(1).Value
        TextBox3.Text = DGV.Rows(e.RowIndex).Cells(2).Value
        TextBox4.Text = DGV.Rows(e.RowIndex).Cells(3).Value
    End Sub


    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Call koneksi()
        Dim edit As String = "update TableJabatan set Nama_Jabatan='" & TextBox2.Text & "',Gaji_Pokok'" & TextBox3.Text & "',Tj_Jabatan='" & TextBox4.Text & "' where Kode_Jabatan='" & Label5.Text & "'"
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