Imports System.Data.SqlClient
Public Class Form1
    Sub Kosong()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox1.Focus()
    End Sub
    Sub Isi()
        TextBox2.Clear()
        TextBox2.Focus()
    End Sub
    Sub TampilJenis()
        da = New SqlDataAdapter("Select * from TBL_Buku", Conn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "TBL_Buku")
        DGV1.DataSource = (ds.Tables("TBL_Buku"))
        DGV1.Refresh()
    End Sub
    Sub AturGrid()
        DGV1.Columns(0).Width = 60
        DGV1.Columns(1).Width = 200
        DGV1.Columns(0).HeaderText = "KODE JENIS"
        DGV1.Columns(1).HeaderText = "NAMA JENIS"
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call Koneksi()
        Call TampilJenis()
        Call Kosong()
        Call AturGrid()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Data Belum Lengkap..!")
            TextBox1.Focus()
            Exit Sub
        Else
            cmd = New SqlCommand("Select * from TBL_Buku" & TextBox1.Text & "", Conn)
            rd = cmd.ExecuteReader
            rd.Read()
            If Not rd.HasRows Then
                Dim Simpan As String = "insert into Jenis(KodeJenis)values " &
                "('" & TextBox1.Text & "','" & TextBox2.Text & "')"
                cmd = New SqlCommand(Simpan, Conn)
                cmd.ExecuteNonQuery()
                MsgBox("Simpan data sukses...!", MsgBoxStyle.Information, "Perhatian")
            End If
            Call TampilJenis()
            Call Kosong()
            TextBox1.Focus()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MsgBox("Kode Jenis belum diisi")
            TextBox1.Focus()
            Exit Sub
        Else
            Dim Ubah As String = "Update Jenis set" &
            "Jenis='" & TextBox2.Text & "' " &
            "where KodeJenis='" & TextBox1.Text & "'"
            cmd = New SqlCommand(Ubah, Conn)
            cmd.ExecuteNonQuery()
            MsgBox("Ubah data sukses..!", MsgBoxStyle.Information, "Perhatian")
            Call TampilJenis()
            Call Kosong()
            TextBox1.Focus()
        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        TextBox2.MaxLength = 50
        If e.KeyChar = Chr(13) Then
            TextBox2.Text = UCase(TextBox2.Text)
        End If
    End Sub

    Private Sub DGV1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV1.CellContentClick
        Dim i As Integer
        i = Me.DGV1.CurrentRow.Index
        With DGV1.Rows.Item(i)
            Me.TextBox1.Text = .Cells(0).Value
            Me.TextBox2.Text = .Cells(1).Value
        End With
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MsgBox("Kode Buku Belum diisi")
            TextBox1.Focus()
            Exit Sub
        Else
            If MessageBox.Show("Yakin akan menghapus Data Jenis" & TextBox1.Text & " ?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                cmd = New SqlCommand("Delete * From Jenis where KodeJenis='" & TextBox1.Text & "'", Conn)
                cmd.ExecuteNonQuery()
                Call Kosong()
                Call TampilJenis()
            Else
                Call Kosong()
            End If
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call Kosong()
    End Sub



    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        TextBox1.MaxLength = 2
        If e.KeyChar = Chr(13) Then
            cmd = New SqlCommand("Select * From Jenis where KodeJenis='" & TextBox1.Text & "'", Conn)
            rd = cmd.ExecuteReader
            rd.Read()
            If rd.HasRows = True Then
                TextBox2.Text = rd.Item(1)
                TextBox2.Focus()
            Else
                Call Isi()
                TextBox2.Focus()
            End If
        End If
    End Sub


    Private Sub TextBox3_Click(sender As Object, e As EventArgs) Handles TextBox3.Click
        cmd = New SqlCommand("Select * From Jenis where KodeJenis like '%" & TextBox3.Text & "%'", Conn)
        rd = cmd.ExecuteReader
        rd.Read()
        If rd.HasRows Then
            da = New SqlDataAdapter("Select * From Jenis where KodeJenis like '%" & TextBox3.Text & "%'", Conn)
            ds = New DataSet
            da.Fill(ds, "Dapat")
            DGV1.DataSource = ds.Tables("Dapat")
            DGV1.ReadOnly = True
        Else
            MsgBox("Data tidak ditemukan")
        End If
    End Sub
End Class
