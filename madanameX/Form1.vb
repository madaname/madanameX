Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'Call upd1punomise()
        ' 任意の間隔で定期的に実行
        ' https://msdn.microsoft.com/ja-jp/library/cyhse5xw(v=vs.90).aspx

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call upd1punomise()
        Me.Close()

    End Sub
End Class
