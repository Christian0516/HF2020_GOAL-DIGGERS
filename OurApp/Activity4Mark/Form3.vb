Public Class Form3
    Dim connect As New OleDb.OleDbConnection
    Dim sql As String
    Dim ds As New DataSet
    Dim da As OleDb.OleDbDataAdapter
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        connect.ConnectionString = "Provider=Microsoft.ace.OLEDB.12.0; data source =|datadirectory|/MARK.accdb"
        ds.Clear()
        Dim input As String
        input = InputBox("Enter NUMBER to search:")
        connect.Open()
        sql = "select * from Table1 where NUMBER='" & input & "'"
        da = New OleDb.OleDbDataAdapter(sql, connect)
        da.Fill(ds, "Table1")
        DataGridView1.DataSource = ds
        DataGridView1.DataMember = "Table1"
        connect.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
        Form6.Show()
    End Sub
End Class