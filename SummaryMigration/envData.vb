Option Explicit On

Friend Class envData
    Private _appPath As String
    Private _appIni As String
    Private _appLog As String
    Private _dbPath As String
    Private _editorPath As String
    Private _csvPath As String
    Private _csvId As String
    Private _csvPass As String
    Private _csvFile As String
    Private _csvDrive As String
    Private _targetCode As List(Of String)
    Private _rootDir As String

    Friend Sub New()
        _appPath = "C:\SummaryMigration"
        _appIni = "\conf\SummaryMigration.ini"
        _appLog = "\log"
        _dbPath = "\conf\SummaryData.sqlite"
        _editorPath = "C:\Windows\System32\notepad.exe"
        _csvPath = ""
        _csvId = ""
        _csvPass = ""
        _csvFile = "*.csv"
        _csvDrive = "S"
        _targetCode = New List(Of String)
        _rootDir = "."
    End Sub

    ''' <summary>
    ''' アプリ配置フォルダ
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
    ''' CSVデータ格納先DB
    ''' </summary>
    ''' <returns></returns>
    Friend ReadOnly Property dbPath() As String
        Get
            Return _dbPath
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
    ''' CSVファイル格納先
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
    ''' CSVファイル格納先アクセスID
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
    ''' CSVファイル格納先アクセスパスワード
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
    ''' CSVファイル名規約
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

    ''' <summary>
    ''' NWドライブ名
    ''' </summary>
    ''' <returns></returns>
    Friend ReadOnly Property csvDrive() As String
        Get
            Return _csvDrive
        End Get
    End Property

    ''' <summary>
    ''' 移行対象の文書コードを設定
    ''' </summary>
    ''' <param name="str"></param>
    Friend Sub setTargetCode(str As String)
        Dim tmp() As String

        tmp = Split(str, ",")
        For i = 0 To tmp.Count - 1
            _targetCode.Add(tmp(i))
        Next
    End Sub

    ''' <summary>
    ''' 移行対象の文書コードリストを取得
    ''' </summary>
    ''' <returns></returns>
    Friend Function getTargetCode() As List(Of String)
        Return _targetCode
    End Function

    ''' <summary>
    ''' 移行先ルートディレクトリ
    ''' </summary>
    ''' <returns></returns>
    Friend Property rootDir As String
        Set(value As String)
            _rootDir = value
        End Set
        Get
            Return _rootDir
        End Get
    End Property
End Class
