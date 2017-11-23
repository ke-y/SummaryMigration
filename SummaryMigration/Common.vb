Option Explicit On

Module Common
    Public Const RESOURCE_CONNECTED = &H1
    Public Const RESOURCETYPE_DISK = &H1
    Public Const RESOURCEUSAGE_CONNECTABLE = &H1
    Public Const RESOURCEDISPLAYTYPE_GENERIC = &H0

    Public Structure NETRESOURCE
        Public dwScope As Integer
        Public dwType As Integer
        Public dwDisplayType As Integer
        Public dwUsage As Integer
        Public lpLocalName As String
        Public lpRemoteName As String
        Public lpComment As String
        Public lpProvider As String
    End Structure

    Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (ByRef lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Integer) As Integer
    Public Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Integer, ByVal fForce As Integer) As Integer

    ''' <summary>
    ''' ネットワークドライブ接続
    ''' </summary>
    ''' <param name="drive">ドライブ名</param>
    ''' <param name="dir">ディレクトリパス</param>
    ''' <param name="user">アクセスユーザ</param>
    ''' <param name="password">パスワード</param>
    ''' <param name="msg">エラーメッセージ</param>
    ''' <returns></returns>
    Public Function loginDir(ByVal drive As String, ByVal dir As String, ByVal user As String, ByVal password As String, ByRef msg As String) As Boolean
        Dim myResource As New NETRESOURCE
        Dim result As Integer

        If Trim(dir) = "" Then
            loginDir = False
        Else
            myResource.dwScope = RESOURCE_CONNECTED
            myResource.dwType = RESOURCETYPE_DISK
            myResource.dwDisplayType = RESOURCEDISPLAYTYPE_GENERIC
            myResource.dwUsage = RESOURCEUSAGE_CONNECTABLE
            myResource.lpComment = Nothing
            myResource.lpLocalName = drive & ":"
            myResource.lpProvider = Nothing
            myResource.lpRemoteName = dir

            result = WNetAddConnection2(myResource, password, user, 0)
            Select Case result
                Case 0  '正常
                    loginDir = True
                Case 5  'アクセス拒否
                    msg = result & " : アクセス拒否されました"
                    loginDir = False
                Case 55 '利用できないリソース
                    msg = result & " : 指定されたフォルダにアクセスできません"
                    loginDir = False
                Case 85 'すでに接続済み
                    loginDir = True
                Case 86 'パスワードが違う
                    msg = result & " : IDまたはパスワードが違います"
                    loginDir = False
                Case 1219 '同一ユーザでの複数接続は不可
                    '接続済みのセッションを使用して再接続
                    result = WNetAddConnection2(myResource, vbNullString, vbNullString, 0)

                    If result = 0 Then
                        loginDir = True
                    Else
                        msg = result
                        loginDir = False
                    End If
                Case Else
                    msg = result
                    loginDir = False
            End Select
        End If

        Return loginDir
    End Function

    ''' <summary>
    ''' ネットワークドライブ切断
    ''' </summary>
    ''' <param name="drive">ドライブ名</param>
    ''' <param name="msg">エラーメッセージ</param>
    ''' <returns></returns>
    Public Function logoffDir(ByVal drive As String, ByRef msg As String) As Boolean
        Dim result As Integer

        logoffDir = False

        '次回ログオン時に再接続しない
        result = WNetCancelConnection2(drive & ":", RESOURCE_CONNECTED, True)

        '開いているファイルやジョブが存在する場合、切断しない
        '存在する場合、エラーになる。（result <> 0）
        'result = WNetCancelConnection2("X:", 0, False)

        Select Case result
            Case 0& '正常
                logoffDir = True
            Case 2250& 'すでに切断済み
                logoffDir = True
            Case Else
                msg = result
                logoffDir = False
        End Select

        Return logoffDir
    End Function

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
