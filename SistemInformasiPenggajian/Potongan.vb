Imports System.Data.SqlClient
Public Class Potongan
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
        TextBox2.Focus()
    End Sub

    Sub DataBaru()
        TextBox2.Clear()
        TextBox2.Focus()
    End Sub
    Sub TampilGrid()
        Call koneksi()
        DA = New SqlDataAdapter("select*from TablePotongan", CONN)
        DS = New DataSet
        DA.Fill(DS)
        DGV.DataSource = DS.Tables(0)
        DGV.ReadOnly = True
    End Sub
    Sub Otomatis()
        CMD = New SqlCommand("select * from TablePotongan order by Kode_Potongan desc", CONN)
        DR = CMD.ExecuteReader
        DR.Read()

        If Not DR.HasRows Then
            Label3.Text = "KP001"
        Else
            Label3.Text = Val(Microsoft.VisualBasic.Mid(DR.Item("Kode_Potongan").ToString, 3, 3)) + 1

            If Len(Label3.Text) = 1 Then
                Label3.Text = "KP00" & Label3.Text & ""
            ElseIf Len(Label3.Text) = 2 Then
                Label3.Text = "KP0" & Label3.Text & ""
            ElseIf Len(Label3.Text) = 3 Then
                Label3.Text = "KP" & Label3.Text & ""
            End If
        End If
    End Sub
    Private Sub label3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar = Chr(13) Then
            Call koneksi()
            CMD = New SqlCommand("select* from TablePotongan where Kode_Potongan='" & Label3.Text & "'", CONN)
            DR = CMD.ExecuteReader
            DR.Read()
            If DR.HasRows Then
                TextBox2.Text = DR.Item("Nama_Potongan")
                TextBox2.Focus()
            Else
                Call DataBaru()
            End If
        End If
    End Sub
    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        TextBox2.MaxLength = 30
        If e.KeyChar = Chr(13) Then
            Button1.Focus()
        End If
    End Sub

    Private Sub Potongan_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Call Otomatis()
    End Sub
    Private Sub Potongan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call koneksi()
        Call TampilGrid()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Label3.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Data Belum Lengkap", MsgBoxStyle.Critical, "WARNING!")
            Exit Sub
        Else
            Call koneksi()
            CMD = New SqlCommand("select * from TablePotongan where Kode_Potongan ='" & Label3.Text & "'", CONN)
            DR = CMD.ExecuteReader
            DR.Read()
            If Not DR.HasRows Then
                Call koneksi()
                Dim simpan As String = "insert into TablePotongan values('" & Label3.Text & "','" & TextBox2.Text & "')"
                CMD = New SqlCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
            Else
                MsgBox("Maaf, Data dengan kode tersebut telah ada",
    MsgBoxStyle.Exclamation, "Peringatan")
                Call DataBaru()
            End If
            Call Otomatis()
            Call DataBaru()
            Call TampilGrid()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Label3.Text = "" Then
            MsgBox("Kode Harus Diisi", MsgBoxStyle.Critical, "WARNING!")
            Exit Sub
        Else
            If MessageBox.Show("Hapus Data ini...?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Call koneksi()
                Dim hapus As String = "delete from TableJabatan where Kode_Jabatan='" & Label3.Text & "'"
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
        Call DataBaru()
        Call koneksi()
        Call Otomatis()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Close()
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        Call koneksi()
        CMD = New SqlCommand("select*from TablePotongan where Nama_Potongan like'%" & TextBox3.Text & "%'", CONN)
        DR = CMD.ExecuteReader
        If DR.HasRows Then
            Call koneksi()
            DA = New SqlDataAdapter("select*from TablePotongan where Nama_Potongan like'%" & TextBox3.Text & "%'", CONN)
            DS = New DataSet
            DA.Fill(DS)
            DGV.DataSource = DS.Tables(0)
        Else
            MsgBox("Nama Potongan Tidak ditemukan", MsgBoxStyle.Critical, "INFORMASI!")
        End If
    End Sub
    Private Sub DGV_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DGV.CellMouseClick
        On Error Resume Next
        Label3.Text = DGV.Rows(e.RowIndex).Cells(0).Value
        TextBox2.Text = DGV.Rows(e.RowIndex).Cells(1).Value
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Call koneksi()
        Dim edit As String = "update TablePotongan set Nama_Potongan='" & TextBox2.Text & "' where Kode_Potongan='" & Label3.Text & "'"
        CMD = New SqlCommand(edit, CONN)
        CMD.ExecuteNonQuery()
        Call DataBaru()
        Call TampilGrid()
        Call Otomatis()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MenuUtama.Show()
    End Sub

    Private Sub DGV_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGV.CellContentClick

    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click

    End Sub
End Class