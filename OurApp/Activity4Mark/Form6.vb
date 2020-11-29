Public Class Form6
    Dim connect As New OleDb.OleDbConnection
    Dim sql As String
    Dim ds As New DataSet
    Dim da As OleDb.OleDbDataAdapter
    Dim inc As Integer
    Dim MaxRows As Integer
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Me.Hide()
        Form3.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If inc <> 0 Then
            inc = 0
            NavigateRecords()
        End If
    End Sub

    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        connect.ConnectionString = "Provider=Microsoft.ace.OLEDB.12.0; data source = |datadirectory|/MARK.accdb"
        connect.Open()
        sql = "select * from Table1"
        da = New OleDb.OleDbDataAdapter(sql, connect)
        da.Fill(ds, "Table1")
        MaxRows = ds.Tables("Table1").Rows.Count
        inc = -1
        connect.Close()
    End Sub
    Private Sub NavigateRecords()
        TextBox1.Text = ds.Tables("Table1").Rows(inc).Item(0)
        TextBox2.Text = ds.Tables("Table1").Rows(inc).Item(1)
        TextBox3.Text = ds.Tables("Table1").Rows(inc).Item(2)
        TextBox4.Text = ds.Tables("Table1").Rows(inc).Item(3)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If inc > 0 Then
            inc = inc - 1
            NavigateRecords()
            If inc = 0 Then
                MsgBox("First Record")

            End If
        ElseIf inc = 0 Then
            MsgBox("No Records Yet")
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If inc <> MaxRows - 1 Then
            inc = inc + 1
            NavigateRecords()
        Else
            MsgBox("No More Rows")

        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If inc <> MaxRows - 1 Then
            inc = MaxRows - 1
            NavigateRecords()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        connect.Open()
        Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand("SELECT * from Table2", connect)
        sql = "insert into Table1 values('" & TextBox1.Text & "' ,'" & TextBox2.Text & "' ,'" & TextBox3.Text & "','" & TextBox4.Text & "')"
        cmd = New OleDb.OleDbCommand(sql, connect)
        cmd.ExecuteNonQuery()
        connect.Close()
        MsgBox("DATA INSERTED")
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        connect.Open()
        Dim cmd2 As OleDb.OleDbCommand = New OleDb.OleDbCommand("Select * from Table1", connect)
        sql = "update Table1 set [NUMBER] = '" & TextBox1.Text & "', [S_NAME] = '" & TextBox2.Text & "', [AGE] = '" & TextBox3.Text & "', [BIRTHDAY] = '" & TextBox4.Text & "' where [NUMBER] = '" & TextBox1.Text & "'"
        cmd2 = New OleDb.OleDbCommand(sql, connect)
        cmd2.ExecuteNonQuery()
        connect.Close()
        MsgBox("Update")
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        connect.Open()
        Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand("Select * from Table1 where NUMBER='" & TextBox1.Text & "'", connect)
        Dim sdr As OleDb.OleDbDataReader = cmd.ExecuteReader
        If (MsgBox("DO YOU WANT OT DELETE THE RECORD?", vbOKCancel) = vbOK) Then
            sql = "Delete * from Table1 where NUMBER='" & TextBox1.Text & "'"
            cmd = New OleDb.OleDbCommand(sql, connect)
            cmd.ExecuteNonQuery()
            MsgBox("Record Deleted")
        Else
            MsgBox("PERMISSION DECLINED")

            connect.Close()
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Form4.Show()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Application.Exit()
    End Sub
End Class