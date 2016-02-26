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


End Module
