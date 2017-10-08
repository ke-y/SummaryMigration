Option Explicit On

Module Common

    ''' <summary>
    ''' ファイルもしくはフォルダが存在するか確認
    ''' </summary>
    ''' <param name="pass">対象データ</param>
    ''' <param name="file">True=ファイル、False=ディレクトリ</param>
    ''' <returns>True/False</returns>
    Friend Function checkExists(ByVal pass As String, ByVal file As Boolean) As Boolean
        If Trim(pass) = "" Then
            checkExists = False
        Else
            If file Then
                checkExists = IO.File.Exists(pass)
            Else
                checkExists = IO.Directory.Exists(pass)
            End If
        End If

        Return checkExists
    End Function

    ''' <summary>
    ''' ログメッセージ出力
    ''' </summary>
    ''' <param name="pass">LOG出力先フォルダ</param>
    ''' <param name="file">LOGファイル名</param>
    ''' <param name="cmt">出力コメント</param>
    Friend Sub putLog(ByVal pass As String, ByVal file As String, ByVal cmt As String)
        Dim sWrite As IO.StreamWriter

        If Not checkExists(pass, False) Then
            IO.Directory.CreateDirectory(pass)
        End If

        If checkExists(pass & "\" & file, True) Then
            'ファイルがすでにある場合は追記する
            sWrite = New IO.StreamWriter(pass & "\" & file, True, Text.Encoding.GetEncoding("Shift_JIS"))
        Else
            sWrite = New IO.StreamWriter(pass & "\" & file, False, Text.Encoding.GetEncoding("Shift_JIS"))
        End If

        sWrite.WriteLine(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & " : " & cmt)
        sWrite.Close()
    End Sub
End Module
