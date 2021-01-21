Imports ExcelLib.Client
Public Class Form1
    Private xls As New ExcelLib.Client 'object untuk menampilkan data ke dalam grid
    Private xls2 As New ExcelLib.Client 'object untuk menampilkan print preview

    Private Sub btnLoad_Click(sender As System.Object, e As System.EventArgs) Handles btnLoad.Click
        Reload()
    End Sub

    Private Sub Reload()
        xls.FileName = "C:\data\obat.xlsx"
        xls.SheetName = "Sheet1"
        xls.Visible = False
        xls.OpenFile()

        xls.LoadToGrid(DataGridView1)
        With DataGridView1
            .RowHeadersVisible = True
            .Columns(0).HeaderText = "No."
            .Columns(1).HeaderText = "Kode Obat"
            .Columns(2).HeaderText = "Nama Obat"
            .Columns(3).HeaderText = "Kemasan"
            .Columns(4).HeaderText = "Volume"
            .Columns(5).HeaderText = "Harga"
            .Columns(6).HeaderText = "Kode Warna"
        End With
    End Sub

    Private Sub btnPreview_Click(sender As System.Object, e As System.EventArgs) Handles btnPreview.Click
        xls2.FileName = "C:\data\obat.xlsx"
        xls2.SheetName = "Sheet1"
        xls2.Visible = True
        xls2.OpenFile()
        xls2.DisplayPreview()
    End Sub

    Private Sub Form1_Disposed(sender As Object, e As System.EventArgs) Handles Me.Disposed
        xls = Nothing
        xls2 = Nothing
    End Sub

End Class



