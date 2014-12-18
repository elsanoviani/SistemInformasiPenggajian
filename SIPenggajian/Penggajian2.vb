Imports System.Data.SqlClient
Public Class Penggajian
    Dim tinggi As Double = 0
    Protected Overrides Sub OnFormClosing(ByVal e As FormClosingEventArgs)
        MyBase.OnFormClosing(e)
        If Not e.Cancel AndAlso e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            MenuUtama.Show()
            Me.Hide()
        End If
    End Sub
    Sub otomatis()
        REM membuat no faktur otomatis
        Dim strtemp As String = ""
        Dim strvalue As String = ""
        Call Koneksi()
        Dim FAK As String = "select * from TableGaji order by Nomor_slip desc"
        CMD = New SqlCommand(FAK, CONN)
        DR = CMD.ExecuteReader
        If DR.HasRows Then
            DR.Read()
            strtemp = Mid(DR.Item("Nomor_slip"), 4, 4)
        Else
            Label1.Text = "DP-0001"
            Exit Sub
        End If
        strvalue = Val(strtemp) + 1
        Label1.Text = "DP-" & Mid("0000", 1, 4 - strvalue.Length) & strvalue
    End Sub

    Sub Bersihkan()
        TextBox1.Text = ""
        Label3.Text = ""
        Label4.Text = ""
        Label5.Text = ""
        Label6.Text = ""
        Label7.Text = ""
        Label8.Text = ""
        Label9.Text = ""
        Label10.Text = ""
        Label11.Text = ""
        Label12.Text = ""
        Label13.Text = ""
        Label14.Text = ""
        Label15.Text = ""
        Label16.Text = ""
        Label17.Text = ""
        DGV.Rows.Clear()
        TextBox1.Focus()
    End Sub
    Sub Hitung()
        Dim hitung As Integer = 0
        For baris As Integer = 0 To DGV.RowCount - 1
            hitung = hitung + DGV.Rows(baris).Cells(2).Value
            Label16.Text = hitung
            Label17.Text = Val(Label15.Text) - hitung
        Next
    End Sub

    Private Sub Penggajian_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Call otomatis()
    End Sub
   

    Private Sub Penggajian_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label2.Text = Format(Now, "dd-MM-yyyy")
        DGV.AlternatingRowsDefaultCellStyle.BackColor = Color.LightBlue
        Call Hitung()

    End Sub


    Private Sub TextBox1_KeyPress1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call koneksi()
            CMD = New SqlCommand("select* from TablePegawai, TableJabatan,TableGolongan where NIP='" & TextBox1.Text & "' And TablePegawai.Golongan=TableGolongan.Golongan and TablePegawai.Kode_Jabatan=TableJabatan.Kode_Jabatan", CONN)
            DR = CMD.ExecuteReader
            DR.Read()
            If DR.HasRows Then
                TextBox1.Text = DR.Item("NIP")
                Label3.Text = DR.Item("Nama_Pegawai")
                Label4.Text = DR.Item("Gaji_Pokok")
                Label5.Text = DR.Item("Tj_Jabatan")
                Label6.Text = DR.Item("Tj_Suami_Istri")
                Label7.Text = DR.Item("Tj_Anak")
                Label8.Text = DR.Item("Tj_Beras")
                Label9.Text = DR.Item("Askes")
                Label10.Text = DR.Item("Nama_Jabatan")
                Label11.Text = DR.Item("Golongan")
                Label12.Text = DR.Item("Status")
                Label13.Text = DR.Item("Jumlah_Anak")
                Dim a As String = ""
                Dim b As String = ""
                a = (Val(Label4.Text) + Val(Label5.Text) + Val(Label6.Text) + Val(Label8.Text) + Val(Label9.Text) + Val(Label14.Text))
                Label14.Text = Label13.Text * Label7.Text
                Label15.Text = Val(a)
                b = Val(a) - Val(Label16.Text)

            Else
                If MessageBox.Show("Kode Potongan Tidak Ditemukan, Ingin Buka Daftar NIP?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    ListPegawai.Show()
                End If
                Call Bersihkan()
            End If
        End If
    End Sub
    
    Private Sub DGV_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGV.CellEndEdit
        If e.ColumnIndex = 0 Then
            Call koneksi()
            CMD = New SqlCommand("Select * from TablePotongan where Kode_Potongan='" & DGV.Rows(e.RowIndex).Cells(0).Value & "'", CONN)
            DR = CMD.ExecuteReader
            DR.Read()
            If DR.HasRows Then
                DGV.Rows(e.RowIndex).Cells(1).Value = DR.Item("Nama_Potongan")
                DGV.Rows(e.RowIndex).Cells(2).Value = 0
                SendKeys.Send("{up}")
                Call Hitung()

            Else
                Try
                    DGV.Rows(e.RowIndex).Cells(0).Value = ""
                    DGV.Rows.RemoveAt(DGV.CurrentCell.RowIndex)
                    SendKeys.Send("{down}")
                    If MessageBox.Show("Kode Potongan Tidak Ditemukan, Ingin Buka Daftar potongan?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                        ListPotongan.Show()
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
                Call Hitung()
            End If
        End If

        
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("Data Belum Lengkap", MsgBoxStyle.Critical, "WARNING!")
            Exit Sub
        Else
            Call koneksi()
            CMD = New SqlCommand("select* from TablePegawai, TableJabatan,TableGolongan where NIP='" & TextBox1.Text & "' And TablePegawai.Golongan=TableGolongan.Golongan and TablePegawai.Kode_Jabatan=TableJabatan.Kode_Jabatan", CONN)
            DR = CMD.ExecuteReader
            DR.Read()

            If Not DR.HasRows Then

            Call koneksi()
            Dim simpan1 As String = "insert into TableGaji values ('" & Label1.Text & "', '" & Label2.Text & "', '" & Label16.Text & "','" & Label17.Text & "', '" & TextBox1.Text & "','" & Label15.Text & "', '" & MenuUtama.Panel1.Text & "')"
            CMD = New SqlCommand(simpan1, CONN)
            CMD.ExecuteNonQuery()

            For baris As Integer = 0 To DGV.RowCount - 2
                Call koneksi()
                Dim simpan2 As String = "insert into DetailGaji values ('" & Label1.Text & "','" & DGV.Rows(baris).Cells(0).Value & "','" & DGV.Rows(baris).Cells(2).Value & "')"
                CMD = New SqlCommand(simpan2, CONN)
                CMD.ExecuteNonQuery()
            Next

            '-------------------------
            tinggi = 0
                With PrintDocument1.PrinterSettings.DefaultPageSettings.Landscape = False
                End With

                '--------------------------
            Else
                MsgBox("Nip ini sudah gajian", MsgBoxStyle.Critical, "WARNING!")
                Exit Sub
            End If
        End If
    End Sub
    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        tinggi += 15
        e.Graphics.DrawString("                  Dinas Pendapatan Daerah", New Drawing.Font("verdana", 8), Brushes.Black, 2, tinggi)
        tinggi += 15
        e.Graphics.DrawString("                         ", New Drawing.Font("verdana", 8), Brushes.Black, 2, tinggi)
        tinggi += 15
        e.Graphics.DrawString("                      ", New Drawing.Font("verdana", 8), Brushes.Black, 2, tinggi)
        tinggi += 25
        e.Graphics.DrawString("===========================", New Drawing.Font("verdana", 8), Brushes.Black, 2, tinggi)
        tinggi += 15
        e.Graphics.DrawString("Nomor Slip: " + Label1.Text + vbTab + "User: " + MenuUtama.Panel3.Text, New Drawing.Font("verdana", 8), Brushes.Black, 2, tinggi)
        tinggi += 15
        e.Graphics.DrawString("-------------------------------------------------", New Drawing.Font("verdana", 8), Brushes.Black, 2, tinggi)
        tinggi += 10
        For row As Integer = 0 To DGV.Rows.Count - 2
            tinggi += 15
            e.Graphics.DrawString(DGV.Rows(row).Cells(1).Value.ToString + vbTab + vbTab + DGV.Rows(row).Cells(2).Value.ToString, New Drawing.Font("verdana", 8), Brushes.Black, 2, tinggi)
        Next

        tinggi += 15
        e.Graphics.DrawString("-------------------------------------------------", New Drawing.Font("verdana", 8), Brushes.Black, 2, tinggi)
        tinggi += 15
        e.Graphics.DrawString("NIP " + vbTab + vbTab + vbTab + TextBox1.Text, New Drawing.Font("verdana", 8), Brushes.Black, 2, tinggi)
        tinggi += 15
        e.Graphics.DrawString("Nama " + vbTab + vbTab + vbTab + Label3.Text, New Drawing.Font("verdana", 8), Brushes.Black, 2, tinggi)
        tinggi += 15
        e.Graphics.DrawString("Gaji Kotor " + vbTab + vbTab + vbTab + Label15.Text, New Drawing.Font("verdana", 8), Brushes.Black, 2, tinggi)
        tinggi += 15
        e.Graphics.DrawString("Potongan " + vbTab + vbTab + vbTab + vbTab + Label16.Text, New Drawing.Font("verdana", 8), Brushes.Black, 2, tinggi)
        tinggi += 15
        e.Graphics.DrawString("-------------------------------------------------", New Drawing.Font("verdana", 8), Brushes.Black, 2, tinggi)
        tinggi += 15
        e.Graphics.DrawString("Gaji Bersih " + vbTab + vbTab + vbTab + Label17.Text, New Drawing.Font("verdana", 8), Brushes.Black, 2, tinggi)
        tinggi += 25
        e.Graphics.DrawString("-------------------------------------------------", New Drawing.Font("verdana", 8), Brushes.Black, 2, tinggi)
        tinggi += 15
        e.Graphics.DrawString("Waktu: " + Date.Now, New Drawing.Font("verdana", 8), Brushes.Black, 2, tinggi)
        tinggi += 15
        e.Graphics.DrawString("===========================", New Drawing.Font("verdana", 8), Brushes.Black, 2, tinggi)
        tinggi += 25
        e.Graphics.DrawString("                    TERIMAKASIH", New Drawing.Font("verdana", 8), Brushes.Black, 2, tinggi)
        
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call Bersihkan()
        Call otomatis()
    End Sub
End Class