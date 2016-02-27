Module CommonProc


    '+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
    ' 処理概要：ウェブページが完全に読み込まれるまで待機する処理　
    ' 使用方法：OpenWebWait("IEオブジェクト")
    '+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

    Public Function OpenWebWait(ByVal objIE As SHDocVw.InternetExplorer) As Boolean

        Try
            '読み込み完了まで待つ
            Do While (objIE.Busy OrElse
                objIE.ReadyState <> SHDocVw.tagREADYSTATE.READYSTATE_COMPLETE)
                '無処理
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(100)
            Loop

            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function


    '+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
    ' 処理概要：指定URLで起動し、読み込みを完了させる処理
    ' 使用方法：createView("IEオブジェクト","表示させたいURLの文字列","IE表示・非表示の値[省略可]")
    '+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

    Public Function createView(ByVal objIE As SHDocVw.InternetExplorer,
                            urlName As String,
                            Optional viewFlg As Boolean = True
                            ) As Boolean
        Try
            '指定したURLのページを表示する
            objIE.Navigate(urlName)
            'IE(InternetExplorer)を表示・非表示
            objIE.Visible = viewFlg
            'IEが完全表示されるまで待機
            Call OpenWebWait(objIE)

            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function


    '+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
    ' 処理概要：指定URLへ移動し、読み込みを完了させる処理
    ' 使用方法：createView("IEオブジェクト","表示させたいURLの文字列","IE表示・非表示の値[省略可]")
    '+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

    Public Function naviPage(ByVal objIE As SHDocVw.InternetExplorer,
                            urlName As String
                            ) As Boolean
        Try
            '指定したURLのページを表示する
            objIE.Navigate(urlName)
            'IEが完全表示されるまで待機
            Call OpenWebWait(objIE)

            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function



    '+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
    ' 処理概要：指定時刻で実行する。
    ' 使用方法：onTimeStart("指定時刻")
    '+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

    Public Function onTimeStart(ByVal objSpecifiedtimeIE As String) As Boolean

        ' 現在の日時を取得します
        Dim dNow As DateTime = System.DateTime.Now

        ' 日付と時刻を取得します
        Dim dDate As DateTime = dNow.Date         ' 現在の日付を取得します
        Dim tTime As TimeSpan = dNow.TimeOfDay    ' 現在の時刻を取得します

        Dim Specifiedtime As String

        ' 本日日付 + 指定時刻
        Dim thday1 As Date = DateTime.ParseExact(dDate & " " & Specifiedtime, "yyyy/MM/dd HH:mm:ss", System.Globalization.DateTimeFormatInfo.InvariantInfo)
        Dim thday2 As Date = Now()

        If thday1 > thday2 Then
            Call upd1punomise()
            Return True
        Else
            Return False
        End If

    End Function



End Module
