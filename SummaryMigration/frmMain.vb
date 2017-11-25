Option Explicit On

Public Class frmMain
    Private env As New envData
    Private csvData As New dbConnect

    ''' <summary>
    ''' 起動時処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "===== Start System =====")

        Text = Application.ProductName & " - MAIN"

        readIni()
        initializeDB()

    End Sub

    ''' <summary>
    ''' サマリ情報を取得
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnGetSummaryInfo_Click(sender As Object, e As EventArgs) Handles btnGetSummaryInfo.Click
        Dim csvfile As List(Of String)
        Dim dtTbl As New DataTable

        csvfile = getCsvFilePath()
        registDB(csvfile)

        If csvData.ProcSelect("select * from SummaryData", dtTbl) Then
            dataview.DataSource = dtTbl
        End If

    End Sub

    ''' <summary>
    ''' iniファイルを読み込む
    ''' </summary>
    Private Sub readIni()
        Dim errList As New List(Of String)
        Dim sRead As IO.StreamReader
        Dim file_pass As String
        Dim str As String
        Dim strTmp() As String
        Dim msg As String = ""

        putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "Read the Ini file")

        file_pass = env.appPath & env.appIni
        If checkExists(file_pass, True) Then
            sRead = New IO.StreamReader(file_pass, System.Text.Encoding.GetEncoding("Shift_JIS"))

            Do
                str = sRead.ReadLine()
                If str Is Nothing Then
                    Exit Do
                End If

                If Trim(str) <> "" Then
                    If str.Substring(0, 1) <> "*" Then
                        strTmp = Split(str, "=")
                        Select Case Trim(strTmp(0))
                            Case "EDITOR"
                                If checkExists(Trim(strTmp(1)), True) Then
                                    env.editorPath = Trim(strTmp(1))

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Set the Editor path")
                                Else
                                    errList.Add(Trim(strTmp(1)) & "が存在しません")
                                End If
                            Case "CSV_PATH"
                                If Trim(strTmp(1)) <> "" Then
                                    env.csvPath = Trim(strTmp(1))

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Set the CSV input path")
                                Else
                                    errList.Add("CSV_PATHが未指定です")
                                End If
                            Case "CSV_ID"
                                If Trim(strTmp(1)) <> "" Then
                                    env.csvId = Trim(strTmp(1))

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Set the CSV access ID")
                                End If
                            Case "CSV_PASS"
                                If Trim(strTmp(1)) <> "" Then
                                    env.csvPass = Trim(strTmp(1))

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Set the CSV access Pass")
                                End If
                            Case "CSV_FILE"
                                If Trim(strTmp(1)) <> "" Then
                                    env.csvFile = Trim(strTmp(1))

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Set the CSV file name pattern")
                                End If
                        End Select
                    End If
                End If
            Loop
            sRead.Close()
        Else
            MsgBox(file_pass & "が存在しません" & vbCrLf & "システムを終了します")

            putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "===== Stop System =====")
            Close()
        End If

        If loginDir(env.csvDrive, env.csvPath, env.csvId, env.csvPass, msg) Then
            putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Access OK " & " [" & env.csvPath & "]")

            If Not logoffDir(env.csvDrive, msg) Then
                putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  CAUTION : Network Dir logoff Error. Win32 API Error Code = " & msg & " [" & env.csvPath & "]")
            End If
        Else
            errList.Add(env.csvPath & "にアクセスできません")

            putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  CAUTION : Network Dir login Error. Win32 API Error Code = " & msg & " [" & env.csvPath & "]")
        End If

        If errList.Count > 0 Then
            For i = 0 To errList.Count - 1
                MsgBox(errList.Item(i))
            Next i
        End If

    End Sub

    ''' <summary>
    ''' データ格納DBの初期化
    ''' </summary>
    Private Sub initializeDB()
        Dim file_pass As String
        Dim sql As New List(Of String)
        Dim count As Integer

        putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "Initialize DB")

        file_pass = env.appPath & env.dbPath
        If checkExists(file_pass, True) Then
            csvData.setPath = file_pass

            sql.Clear()
            sql.Add("delete from SummaryData")
            If csvData.ProcUpdate(sql, count) Then
                putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Success to initialize DB")
            Else
                putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  CAUTION : Failed to initialize DB")
            End If
        Else
            MsgBox(file_pass & "が存在しません" & vbCrLf & "システムを終了します")

            putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "===== Stop System =====")
            Close()
        End If

    End Sub

    ''' <summary>
    ''' CSVファイルのパスを取得
    ''' </summary>
    ''' <returns></returns>
    Private Function getCsvFilePath() As List(Of String)
        Dim csvlist As New List(Of String)
        Dim dirInfo As IO.DirectoryInfo
        Dim allFile As IO.FileInfo()
        Dim msg As String = ""

        putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "Search CSV File")

        csvlist.Clear()
        If loginDir(env.csvDrive, env.csvPath, env.csvId, env.csvPass, msg) Then
            dirInfo = New IO.DirectoryInfo(env.csvPath)
            allFile = Nothing

            allFile = dirInfo.GetFiles(env.csvFile, IO.SearchOption.AllDirectories)
            For Each f As IO.FileInfo In allFile
                csvlist.Add(f.FullName)
            Next

            If csvlist.Count > 0 Then
                putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Get the CSV File count = " & csvlist.Count)
            Else
                putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  CAUTION : CSV File is Not Found")
            End If

            If Not logoffDir(env.csvDrive, msg) Then
                putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  CAUTION : Network Dir logoff Error. Win32 API Error Code = " & msg & " [" & env.csvPath & "]")
            End If
        Else
            putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  CAUTION : Network Dir login Error. Win32 API Error Code = " & msg & " [" & env.csvPath & "]")
        End If

        Return csvlist
    End Function

    ''' <summary>
    ''' CSVファイルの内容をDBに登録
    ''' </summary>
    ''' <param name="csvlist"></param>
    Private Sub registDB(csvlist As List(Of String))
        Dim sRead As IO.StreamReader
        Dim strline As String
        Dim strSql As New List(Of String)
        Dim index As Integer
        Dim msg As String = ""

        putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "Regist File Data with DB")

        If loginDir(env.csvDrive, env.csvPath, env.csvId, env.csvPass, msg) Then
            For index = 0 To csvlist.Count - 1
                sRead = New IO.StreamReader(csvlist(index), System.Text.Encoding.GetEncoding("Shift_JIS"))

                Do
                    strline = sRead.ReadLine()
                    If strline Is Nothing Then
                        Exit Do
                    End If

                    strSql.Add("insert into SummaryData values (" & strline & ", "" "", "" "")")
                Loop
            Next

            If strSql.Count > 0 Then
                If csvData.ProcUpdate(strSql, index) Then
                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Insert data = " & index)
                Else
                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  CAUTION : Failed to insert DB")
                End If
            Else
                putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  CAUTION : Data is not Found")
            End If

            If Not logoffDir(env.csvDrive, msg) Then
                putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  CAUTION : Network Dir logoff Error. Win32 API Error Code = " & msg & " [" & env.csvPath & "]")
            End If
        Else
            putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  CAUTION : Network Dir login Error. Win32 API Error Code = " & msg & " [" & env.csvPath & "]")
        End If

    End Sub

End Class
