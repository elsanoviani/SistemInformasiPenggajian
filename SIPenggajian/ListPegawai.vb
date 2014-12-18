Imports System.Data.SqlClient

Public Class ListPegawai

    Sub TampilGrid()
        Call koneksi()
        DA = New SqlDataAdapter("select NIP, Nama_Pegawai from TablePegawai", CONN)
        DS = New DataSet
        DA.Fill(DS)
        DGV.DataSource = DS.Tables(0)
        DGV.ReadOnly = True
    End Sub

    Private Sub ListPegawai_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call TampilGrid()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Call koneksi()
        CMD = New SqlCommand("select NIP, Nama_Pegawai from TablePegawai where Nama_Pegawai like'%" & TextBox1.Text & "%'", CONN)
        DR = CMD.ExecuteReader
        If DR.HasRows Then
            Call koneksi()
            DA = New SqlDataAdapter("select NIP, Nama_Pegawai from TablePegawai where Nama_Pegawai like'%" & TextBox1.Text & "%'", CONN)
            DS = New DataSet
            DA.Fill(DS)
            DGV.DataSource = DS.Tables(0)
        Else
            MsgBox("Nama Pegawai Tidak ditemukan", MsgBoxStyle.Critical, "INFORMASI!")
        End If
    End Sub
End Class