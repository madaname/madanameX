Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        ' Call upd1punomise()
        ' 任意の間隔で定期的に実行したい。
        ' https://msdn.microsoft.com/ja-jp/library/cyhse5xw(v=vs.90).aspx

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call onTimeStart("23:59:00")
    End Sub

    Public Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        'Dim thday1 As Date = DateTime.ParseExact("2016/02/27 23:00:00", "yyyy/MM/dd HH:mm:ss", System.Globalization.DateTimeFormatInfo.InvariantInfo)
        'Dim thday2 As Date

        'thday2 = Now()
        'MsgBox(thday2)
        'Me.Label2.Text = thday2

        'thday1 = Me.TextBox1.Text

        'If thday1 > thday2 Then
        '    Me.Label2.Text = "現在時刻 " & thday2
        'Else
        '    Me.Label2.Text = "過ぎました"
        'End If

    End Sub


End Class
