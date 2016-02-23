Public Class Form1

    Private ie As SHDocVw.InternetExplorer

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim form As mshtml.HTMLFormElement ' フォーム送信用
        Dim TITLE As String
        Dim CONTENT As String

        TITLE = "激アツ☆まだ舐めたくて学園イベント情報♪♪♪"
        CONTENT = "超居残りイベントです!!!" & vbCrLf & vbCrLf & "「イベントを見た」と伝えると全コース" & vbCrLf & vbCrLf & "+15分or￥1000割引とさせていただきます!!" & vbCrLf & vbCrLf & "さらにフリーのお客様にも \1000割引させて頂きます。" & vbCrLf & vbCrLf & "60 分16500円 交通費でお遊び頂けます｡" & vbCrLf & vbCrLf & "大変お得なイベントになります(*^^)v" & vbCrLf & vbCrLf & "※ご予約時にお伝えいただた方のみ､有効とさせていただきますので､予めご了承ください｡ "


        'IEの起動
        ie = CreateObject("InternetExplorer.Application") 'オブジェクトを作成
        ie.Navigate("http://deriheru-1m.com/admin/login")   '指定URLで起動
        ie.Visible = True    'IEを表示

        Call OpenWebWait()　     'ieが完全表示されるまで待機

        ' ■■ログイン■■

        ' 1分間のログインIDを入力
        ie.Document.getElementsByName("login_email")(0).Value = "nametakute@gmail.com"
        ' 1分間のパスワードを入力
        ie.Document.getElementsByName("pass")(0).Value = "165116511651"

        form = ie.Document.forms(0)
        form.submit()
        Call OpenWebWait()           'ieが完全表示されるまで待機

        '■■お店速報へ移動■■
        ie.Navigate("http://deriheru-1m.com/admin/flash/")
        Call OpenWebWait()      'ieが完全表示されるまで待機

        'タイトルと内容を初期化と記入
        ie.Document.getElementById("str1").Value = ""
        ie.Document.getElementById("str1").Value = TITLE
        ie.Document.getElementById("str2").Value = ""
        ie.Document.getElementById("str2").Value = CONTENT

        '1秒間 ボーっとする
        System.Threading.Thread.Sleep(1000)

        ' ■■お店速報を投稿する■■
        DirectCast(ie.Document, mshtml.HTMLDocument).forms(0).submit()

        ie.Quit()

        MsgBox("OK")

    End Sub

    Public Function OpenWebWait() As Boolean

        Try
            '読み込み完了まで待つ
            Do While (ie.Busy OrElse
                ie.ReadyState <> SHDocVw.tagREADYSTATE.READYSTATE_COMPLETE)

                '無処理
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(100)
            Loop

            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function



End Class
