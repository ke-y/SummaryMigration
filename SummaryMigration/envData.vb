Option Explicit On

Friend Class envData
    Private _appPath As String
    Private _appIni As String
    Private _appLog As String
    Private _editorPath As String
    Private _csvPath As String
    Private _csvId As String
    Private _csvPass As String
    Private _csvFile As String

    Friend Sub New()
        _appPath = "C:\SummaryMigration"
        _appIni = "\conf\SummaryMigration.ini"
        _appLog = "\log"
        _editorPath = "C:\Windows\System32\notepad.exe"
        _csvPath = ""
        _csvId = ""
        _csvPass = ""
        _csvFile = "*.csv"
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

    ''' <summary>
    ''' CSV格納先
    ''' </summary>
    ''' <returns></returns>
    Friend Property csvPath() As String
        Set(ByVal value As String)
            _csvPath = value
        End Set
        Get
            Return _csvPath
        End Get
    End Property

    ''' <summary>
    ''' CSV格納先
    ''' </summary>
    ''' <returns></returns>
    Friend Property csvId() As String
        Set(ByVal value As String)
            _csvId = value
        End Set
        Get
            Return _csvId
        End Get
    End Property

    ''' <summary>
    ''' CSV格納先
    ''' </summary>
    ''' <returns></returns>
    Friend Property csvPass() As String
        Set(ByVal value As String)
            _csvPass = value
        End Set
        Get
            Return _csvPass
        End Get
    End Property

    ''' <summary>
    ''' CSV格納先
    ''' </summary>
    ''' <returns></returns>
    Friend Property csvFile() As String
        Set(ByVal value As String)
            _csvFile = value
        End Set
        Get
            Return _csvFile
        End Get
    End Property

End Class
