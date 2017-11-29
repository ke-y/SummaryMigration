Option Explicit On

Friend Class envData
    Private _appPath As String
    Private _appIni As String
    Private _appLog As String
    Private _dbPath As String
    Private _editorPath As String
    Private _summaryPath As String
    Private _loginId As String
    Private _loginPass As String
    Private _csvFile As String
    Private _copyFile As String
    Private _summaryDrive As String
    Private _targetCode As List(Of String)
    Private _rootDir As String
    Private _pidLen As Integer
    Private _pidChr As Char

    Friend Sub New()
        _appPath = "C:\SummaryMigration"
        _appIni = "\conf\SummaryMigration.ini"
        _appLog = "\log"
        _dbPath = "\conf\SummaryData.sqlite"
        _editorPath = "C:\Windows\System32\notepad.exe"
        _summaryPath = ""
        _loginId = ""
        _loginPass = ""
        _csvFile = "*.csv"
        _copyFile = "*.pdf"
        _summaryDrive = "S"
        _targetCode = New List(Of String)
        _rootDir = "."
        _pidLen = 10
        _pidChr = "0"
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
    Friend Property summaryPath() As String
        Set(ByVal value As String)
            _summaryPath = value
        End Set
        Get
            Return _summaryPath
        End Get
    End Property

    ''' <summary>
    ''' CSVファイル格納先アクセスID
    ''' </summary>
    ''' <returns></returns>
    Friend Property loginId() As String
        Set(ByVal value As String)
            _loginId = value
        End Set
        Get
            Return _loginId
        End Get
    End Property

    ''' <summary>
    ''' CSVファイル格納先アクセスパスワード
    ''' </summary>
    ''' <returns></returns>
    Friend Property loginPass() As String
        Set(ByVal value As String)
            _loginPass = value
        End Set
        Get
            Return _loginPass
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
    ''' コピー対象のファイル名規約
    ''' </summary>
    ''' <returns></returns>
    Friend Property copyFile() As String
        Set(ByVal value As String)
            _copyFile = value
        End Set
        Get
            Return _copyFile
        End Get
    End Property

    ''' <summary>
    ''' NWドライブ名
    ''' </summary>
    ''' <returns></returns>
    Friend ReadOnly Property summaryDrive() As String
        Get
            Return _summaryDrive
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

    ''' <summary>
    ''' 移行後の患者ID桁数
    ''' </summary>
    ''' <returns></returns>
    Friend ReadOnly Property pidLen() As Integer
        Get
            Return _pidLen
        End Get
    End Property

    ''' <summary>
    ''' 患者IDの桁埋め文字
    ''' </summary>
    ''' <returns></returns>
    Friend ReadOnly Property pidChr() As Char
        Get
            Return _pidChr
        End Get
    End Property

End Class
