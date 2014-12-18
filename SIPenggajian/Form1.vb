Imports System.Data.SqlClient
Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.ComboBox1.Items.Add("Januari")
        Me.ComboBox1.Items.Add("Februari")
        Me.ComboBox1.Items.Add("Maret")
        Me.ComboBox1.Items.Add("April")
        Me.ComboBox1.Items.Add("Mei")
        Me.ComboBox1.Items.Add("Juni")
        Me.ComboBox1.Items.Add("Juli")
        Me.ComboBox1.Items.Add("Agustus")
        Me.ComboBox1.Items.Add("September")
        Me.ComboBox1.Items.Add("Oktober")
        Me.ComboBox1.Items.Add("November")
        Me.ComboBox1.Items.Add("Desember")

        Me.ComboBox2.Items.Add("2014")
        Me.ComboBox2.Items.Add("2015")
        Me.ComboBox2.Items.Add("2016")
        Me.ComboBox2.Items.Add("2017")
        Me.ComboBox2.Items.Add("2018")
        Me.ComboBox2.Items.Add("2019")

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim int_bulan As Integer

        Try

            Select Case ComboBox1.Text

                Case "Januari"

                    int_bulan = 1

                Case "Februari"

                    int_bulan = 2

                Case "Maret"

                    int_bulan = 3

                Case "April"

                    int_bulan = 4

                Case "Mei"

                    int_bulan = 5

                Case "Juni"

                    int_bulan = 6

                Case "Juli"

                    int_bulan = 7

                Case "Agustus"

                    int_bulan = 8

                Case "September"

                    int_bulan = 9

                Case "Oktober"

                    int_bulan = 10

                Case "November"

                    int_bulan = 11

                Case "Desember"

                    int_bulan = 12

            End Select

            If ComboBox1.Text = "" Or ComboBox2.Text = "" Then

                MessageBox.Show("Isikan dengan Lengkap")

            Else

                Laporan.CrystalReportViewer1.SelectionFormula() = "Month({TableGaji.Tanggal}) =" & Val(int_bulan) & " and Year({TableGaji.Tanggal}) =" &
                Val(ComboBox2.Text)

                Laporan.CrystalReportViewer1.RefreshReport()

                Laporan.WindowState = FormWindowState.Maximized

                Laporan.Show()

            End If

        Catch ex As Exception

            MessageBox.Show("Report Error", "Form Filter Report", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
End Class
   