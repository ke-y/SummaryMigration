Option Explicit On

Friend Class envData
    Private _appPath As String
    Private _appIni As String
    Private _appLog As String
    Private _editorPath As String

    Friend Sub New()
        _appPath = "C:\SummaryMigration"
        _appIni = "\conf\SummaryMigration.ini"
        _appLog = "\log"
        _editorPath = "C:\Windows\System32\notepad.exe"
    End Sub

    ''' <summary>
    ''' アプリ配置先フォルダ
    ''' </summary>
    ''' <returns></returns>
    Friend ReadOnly Property appPath() As String
        Get
            Return _appPath
        End Get
    End Property

    ''' <summary>
    ''' INIファイル名
    ''' </summary>
    ''' <returns></returns>
    Friend ReadOnly Property appIni() As String
        Get
            Return _appIni
        End Get
    End Property

    ''' <summary>
    ''' ログ出力フォルダ
    ''' </summary>
    ''' <returns></returns>
    Friend ReadOnly Property appLog() As String
        Get
            Return _appLog
        End Get
    End Property

    ''' <summary>
    ''' エディタ
    ''' </summary>
    ''' <returns></returns>
    Friend Property editorPath() As String
        Set(ByVal value As String)
            _editorPath = value
        End Set
        Get
            Return _editorPath
        End Get
    End Property

End Class
