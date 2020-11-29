Public Class Form1
    Dim connect As New OleDb.OleDbConnection
    Dim attempt As Integer = 0
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        connect.ConnectionString = "Provider=Microsoft.ace.OLEDB.12.0; data source = |datadirectory|/MARK.accdb"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Username, Password As String
        Username = TextBox1.Text
        Password = TextBox2.Text
        If Username = "MARK" And Password = "123" Then
            MsgBox("ACESS GRANTED!")
            Me.Hide()
            Form6.Show()
        ElseIf attempt = 3 Then
            MsgBox("CRITICAL WARNING!", MsgBoxStyle.Critical, "WARNING")
            Close()
        Else
            MsgBox("USERNAME AND PASSWORD DENIED! " & attempt & "of 3.")
            attempt = attempt + 1


        End If
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class
