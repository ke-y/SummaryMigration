Option Explicit On

Friend Class dbConnect

    Private _conn As SQLiteConnection
    Private _cmd As SQLiteCommand
    Private _adp As SQLiteDataAdapter

    Friend Sub New()
        _conn = New SQLiteConnection
        _cmd = New SQLiteCommand
        _adp = New SQLiteDataAdapter
    End Sub

    ''' <summary>
    ''' 接続先DBを指定
    ''' </summary>
    Friend WriteOnly Property setPath() As String
        Set(ByVal value As String)
            _conn.ConnectionString = "DataSource=" & Trim(value)
        End Set
    End Property

    ''' <summary>
    ''' Open処理
    ''' </summary>
    ''' <returns></returns>
    Private Function Open() As Boolean
        Try
            _cmd = _conn.CreateCommand
            _conn.Open()

            Return True
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try

        Return False
    End Function

    ''' <summary>
    ''' Close処理
    ''' </summary>
    ''' <returns></returns>
    Private Function Close() As Boolean
        Try
            If _conn.State = ConnectionState.Open Then _conn.Close()

            Return True
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try

        Return False
    End Function

    ''' <summary>
    ''' SQL実行(INSERT/UPDATE/DELETE)
    ''' </summary>
    ''' <param name="strSql"></param>
    ''' <returns></returns>
    Private Function ExecuteNonQuery(strSql As String) As Integer
        Dim count As Integer = -1

        Try
            If _conn.State = ConnectionState.Open Then
                _cmd.CommandText = strSql
                count = _cmd.ExecuteNonQuery
            End If
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try

        Return count
    End Function

    ''' <summary>
    ''' SQL実行(SELECT)
    ''' </summary>
    ''' <param name="strSQL"></param>
    ''' <returns></returns>
    Private Function ExecuteQuery(strSQL As String) As DataTable
        Dim adapter = New SQLiteDataAdapter()
        Dim dtTbl As New DataTable()

        Try
            If _conn.State <> ConnectionState.Open Then Return dtTbl

            adapter = New SQLiteDataAdapter(strSQL, _conn)
            adapter.Fill(dtTbl)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try

        Return dtTbl
    End Function

    ''' <summary>
    ''' トランザクション開始処理
    ''' </summary>
    ''' <returns></returns>
    Private Function BeginTransaction() As Boolean
        Try
            If Not Open() Then Return False
            If _conn.State = ConnectionState.Closed Then Return False

            _cmd.Transaction = _conn.BeginTransaction()

            Return True
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try

        Return False
    End Function

    ''' <summary>
    ''' ロールバック処理
    ''' </summary>
    ''' <returns></returns>
    Private Function Rollback() As Boolean
        Try
            _cmd.Transaction.Rollback()

            Return True
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try

        Return False
    End Function

    ''' <summary>
    ''' トランザクション終了処理
    ''' </summary>
    ''' <returns></returns>
    Private Function EndTransaction() As Boolean
        Try
            _cmd.Transaction.Commit()
            Close()

            Return True
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try

        Return False
    End Function

    ''' <summary>
    ''' Insert/Update/Delete操作
    ''' </summary>
    ''' <param name="strSql">SQL内容</param>
    ''' <param name="count">処理件数</param>
    ''' <returns></returns>
    Friend Function ProcUpdate(strSql As List(Of String), ByRef count As Integer) As Boolean
        Dim result As Boolean = True
        Dim index As Integer
        Dim tmpnum As Integer

        If Not BeginTransaction() Then Return False

        Try
            count = 0
            For index = 0 To strSql.Count - 1
                tmpnum = ExecuteNonQuery(strSql(index))
                If tmpnum <> -1 Then
                    count = count + tmpnum
                End If
            Next
        Catch ex As Exception
            Rollback()
            result = False
        End Try

        If Not EndTransaction() Then result = False

        Return result
    End Function

    ''' <summary>
    ''' Select操作
    ''' </summary>
    ''' <param name="strSql">SQL内容</param>
    ''' <param name="dtTbl">Select結果</param>
    ''' <returns></returns>
    Friend Function ProcSelect(strSql As String, ByRef dtTbl As DataTable) As Boolean
        Dim result As Boolean = True

        If Not BeginTransaction() Then Return False

        Try
            dtTbl = ExecuteQuery(strSql)
        Catch ex As Exception
            Rollback()
            result = False
        End Try

        If Not EndTransaction() Then result = False

        Return result
    End Function

End Class
