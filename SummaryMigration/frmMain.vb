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

            putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "Display DB Insert Result")
        End If

    End Sub

    ''' <summary>
    ''' iniファイルを読み込む
    ''' </summary>
    Private Sub readIni()
        Dim errList As New List(Of String)
        Dim sRead As IO.StreamReader
        Dim file_pass As String
        Dim strLine As String
        Dim strLineAttr() As String
        Dim msg As String = ""

        putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "Read the Ini file")

        file_pass = env.appPath & env.appIni
        If checkExists(file_pass, True) Then
            sRead = New IO.StreamReader(file_pass, System.Text.Encoding.GetEncoding("Shift_JIS"))

            Do
                strLine = sRead.ReadLine()
                If strLine Is Nothing Then
                    Exit Do
                End If

                If Trim(strLine) <> "" Then
                    If strLine.Substring(0, 1) <> "*" Then
                        strLineAttr = Split(strLine, "=")
                        Select Case Trim(strLineAttr(0))
                            Case "EDITOR"
                                If checkExists(Trim(strLineAttr(1)), True) Then
                                    env.editorPath = Trim(strLineAttr(1))

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Set the Editor path")
                                Else
                                    errList.Add(Trim(strLineAttr(1)) & "が存在しません")
                                End If
                            Case "CSV_PATH"
                                If Trim(strLineAttr(1)) <> "" Then
                                    env.csvPath = Trim(strLineAttr(1))

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Set the CSV input path")
                                Else
                                    errList.Add("CSV_PATHが未指定です")
                                End If
                            Case "CSV_ID"
                                If Trim(strLineAttr(1)) <> "" Then
                                    env.csvId = Trim(strLineAttr(1))

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Set the CSV access ID")
                                End If
                            Case "CSV_PASS"
                                If Trim(strLineAttr(1)) <> "" Then
                                    env.csvPass = Trim(strLineAttr(1))

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Set the CSV access Pass")
                                End If
                            Case "CSV_FILE"
                                If Trim(strLineAttr(1)) <> "" Then
                                    env.csvFile = Trim(strLineAttr(1))

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Set the CSV file name pattern")
                                End If
                            Case "TARGET"
                                If Trim(strLineAttr(1)) <> "" Then
                                    env.setTargetCode(Trim(strLineAttr(1)))

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Set the target DocCode")
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

        If loginDir(env.csvDrive, env.csvPath, env.csvId, env.csvPass, msg) Then
            dirInfo = New IO.DirectoryInfo(env.csvPath)
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
        Dim outputFlg As String = ""
        Dim newName As String = ""
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

                    outputFlg = " "
                    newName = " "
                    checkMigration(strline, outputFlg, newName)
                    strSql.Add("insert into SummaryData values (" & strline & ", " & Chr(34) & outputFlg & Chr(34) & ", " & Chr(34) & newName & Chr(34) & ")")
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

    ''' <summary>
    ''' 移行対象ファイルかチェックして、新しいファイル名を付与
    ''' </summary>
    ''' <param name="strLine"></param>
    ''' <param name="outputFlg"></param>
    ''' <param name="newName"></param>
    Private Sub checkMigration(strLine As String, ByRef outputFlg As String, ByRef newName As String)
        Dim strLineAttr() As String
        Dim strTmp() As String
        Dim docCode As String
        Dim patId As String
        Dim docDate As String
        Dim orderNo As String
        Dim transactionDate As String
        Dim deptCode As String

        strLineAttr = Split(strLine, ",")
        docCode = strLineAttr(3).Replace(Chr(34), "")
        If env.getTargetCode().IndexOf(docCode) <> -1 Then
            outputFlg = "C"

            strTmp = Split(strLineAttr(1), "_")
            patId = strLineAttr(2).Replace(Chr(34), "")
            docDate = strTmp(1)
            orderNo = strTmp(6).Replace(Chr(34), "").Replace(".PDF", "")
            transactionDate = strLineAttr(11).Replace(Chr(34), "")
            deptCode = strLineAttr(5).Replace(Chr(34), "")

            newName = patId & "_" & docDate & "_SMR-02_" & orderNo & "_" & transactionDate & "_" & deptCode & "_1.pdf"
        End If
    End Sub
End Class
