Public Class Form4
    Dim connect As New OleDb.OleDbConnection
    Dim sql As String
    Dim ds As New DataSet
    Dim da As OleDb.OleDbDataAdapter
    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
        Form6.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        connect.ConnectionString = "Provider=Microsoft.ace.OLEDB.12.0; data source=|datadirectory|/MARK.accdb"
        connect.Open()
        da = New OleDb.OleDbDataAdapter("select * from Table1", connect)
        da.Fill(ds, "Table1")
        Dim dv As New DataView(ds.Tables("Table1"))
        DataGridView1.DataSource = dv
        connect.Close()
    End Sub
End Class