Option Explicit On

Public Class frmMain
    Private env As New envData
    Private csvData As New dbConnect
    Private sfd As New SaveFileDialog

    ''' <summary>
    ''' 起動時処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "===== Start System =====")

        Text = Application.ProductName & " - MAIN"
        ControlBox = False
        dataview.AllowUserToAddRows = False
        dataview.ReadOnly = True

        readIni()
        initializeDB()

        btnOutput.Enabled = False
        btnMigration.Enabled = False
        status.Text = "Status:未実行"

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

            btnOutput.Enabled = True
            btnMigration.Enabled = True
            status.Text = "Status:サマリ情報取得済み　テーブル件数=" & dtTbl.Rows.Count

            putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "Display DB Insert Result")
        End If

    End Sub

    ''' <summary>
    ''' 表示内容をCSV出力
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnOutput_Click(sender As Object, e As EventArgs) Handles btnOutput.Click
        Dim file_pass As String = ""

        putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "Save table data in CSV ")

        btnGetSummaryInfo.Enabled = False
        btnOutput.Enabled = False
        btnMigration.Enabled = False

        sfd.Filter = "CSVファイル | *.csv"
        If saveCsv(dataview, file_pass) Then
            putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  output CSV = [" & file_pass & "]")
        Else
            putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  CAUTION : Failed to output CSV")
        End If

        btnGetSummaryInfo.Enabled = True
        btnOutput.Enabled = True
        btnMigration.Enabled = True

    End Sub

    ''' <summary>
    ''' サマリファイルを移行
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnMigration_Click(sender As Object, e As EventArgs) Handles btnMigration.Click
        Dim copyNum As Integer = 0
        Dim checkNum As Integer = 0
        Dim dtTbl As New DataTable

        putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "SummaryFile Migration Start")

        btnGetSummaryInfo.Enabled = False
        btnOutput.Enabled = False
        status.Text = "Status:サマリ移行処理中"

        copyNum = copySummaryFile()
        checkNum = checkSummaryFile()

        If csvData.ProcSelect("select * from SummaryData", dtTbl) Then
            dataview.DataSource = Nothing
            dataview.DataSource = dtTbl
        End If

        btnGetSummaryInfo.Enabled = True
        btnOutput.Enabled = True
        status.Text = "Status:サマリ移行完了  テーブル件数=" & dtTbl.Rows.Count & "　移行件数=" & checkNum & "件　エラー件数=" & copyNum - checkNum & "件"
    End Sub

    ''' <summary>
    ''' iniファイルを開く
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub menuOpenIni_Click(sender As Object, e As EventArgs) Handles menuOpenIni.Click
        openIni()
    End Sub

    ''' <summary>
    ''' iniファイルの内容を反映する
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub menuReadIni_Click(sender As Object, e As EventArgs) Handles menuReadIni.Click
        readIni()
    End Sub

    ''' <summary>
    ''' ログファイルを開く
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub menuOpenLog_Click(sender As Object, e As EventArgs) Handles menuOpenLog.Click
        openLog()
    End Sub

    ''' <summary>
    ''' 格納先ストレージを開く
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub menuOpenStrage_Click(sender As Object, e As EventArgs) Handles menuOpenStrage.Click
        openStrage()
    End Sub

    ''' <summary>
    ''' 終了
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "===== End System =====")

        Close()
    End Sub

    ''' <summary>
    ''' iniファイルを読み込む
    ''' </summary>
    Private Sub readIni()
        Dim errList As New List(Of String)
        Dim sRead As IO.StreamReader
        Dim file_pass As String = ""
        Dim strLine As String = ""
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
                            Case "SUMMARY_PATH"
                                If Trim(strLineAttr(1)) <> "" Then
                                    env.summaryPath = Trim(strLineAttr(1))

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Set the CSV input path")
                                Else
                                    errList.Add("SUMMARY_PATHが未指定です")
                                End If
                            Case "LOGIN_ID"
                                If Trim(strLineAttr(1)) <> "" Then
                                    env.loginId = Trim(strLineAttr(1))

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Set the CSV access ID")
                                End If
                            Case "LOGIN_PASS"
                                If Trim(strLineAttr(1)) <> "" Then
                                    env.loginPass = Trim(strLineAttr(1))

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Set the CSV access Pass")
                                End If
                            Case "CSV_FILE"
                                If Trim(strLineAttr(1)) <> "" Then
                                    env.csvFile = Trim(strLineAttr(1))

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Set the CSV file name pattern")
                                End If
                            Case "COPY_FILE"
                                If Trim(strLineAttr(1)) <> "" Then
                                    env.copyFile = Trim(strLineAttr(1))

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Set the Copy file name pattern")
                                End If
                            Case "TARGET"
                                If Trim(strLineAttr(1)) <> "" Then
                                    env.setTargetCode(Trim(strLineAttr(1)))

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Set the Target DocCode")
                                End If
                            Case "FROM"
                                If Trim(strLineAttr(1)) <> "" And checkFormat(Trim(strLineAttr(1)), "\d{8}") Then
                                    env.fromDay = Trim(strLineAttr(1))

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Set the Search Start day")
                                End If
                            Case "TO"
                                If Trim(strLineAttr(1)) <> "" And checkFormat(Trim(strLineAttr(1)), "\d{8}") Then
                                    env.toDay = Trim(strLineAttr(1))

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Set the Search End day")
                                End If
                            Case "ROOTDIR"
                                If Trim(strLineAttr(1)) <> "" Then
                                    If checkExists(Trim(strLineAttr(1)), False) Then
                                        env.rootDir = Trim(strLineAttr(1))

                                        putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Set the Root Directory")
                                    Else
                                        errList.Add(Trim(strLineAttr(1)) & "が存在しません")
                                    End If
                                Else
                                    errList.Add("ROOTDIRが未指定です")
                                End If
                            Case "DOCCODE"
                                If Trim(strLineAttr(1)) <> "" Then
                                    env.newDocCode = Trim(strLineAttr(1))

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Set the New DocCode")
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

        If env.summaryPath <> "" Then
            If loginDir(env.summaryDrive, env.summaryPath, env.loginId, env.loginPass, msg) Then
                putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Access OK " & " [" & env.summaryPath & "]")

                If Not logoffDir(env.summaryDrive, msg) Then
                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  CAUTION : Network Dir logoff Error. Win32 API Error Code = " & msg & " [" & env.summaryPath & "]")
                End If
            Else
                errList.Add(env.summaryPath & "にアクセスできません")

                putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  CAUTION : Network Dir login Error. Win32 API Error Code = " & msg & " [" & env.summaryPath & "]")
            End If
        End If

        If errList.Count > 0 Then
            For i = 0 To errList.Count - 1
                MsgBox(errList.Item(i))
            Next i
        End If

    End Sub

    ''' <summary>
    ''' iniファイルを開く
    ''' </summary>
    Private Sub openIni()
        Dim file_pass As String

        putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "Open the Ini file")

        file_pass = env.appPath & env.appIni
        If checkExists(file_pass, True) Then
            Shell(env.editorPath & " " & file_pass, vbNormalFocus)
        Else
            MsgBox(file_pass & "が存在しません" & vbCrLf & "システムを終了します")

            putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "===== Stop System =====")
            Close()
        End If

    End Sub

    ''' <summary>
    ''' logファイルを開く
    ''' </summary>
    Private Sub openLog()
        Dim file_pass As String

        putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "Open the Log file")

        file_pass = env.appPath & "\" & env.appLog & "\" & My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log"
        If checkExists(file_pass, True) Then
            Shell(env.editorPath & " " & file_pass, vbNormalFocus)
        End If

    End Sub

    ''' <summary>
    ''' 格納先ストレージを開く
    ''' </summary>
    Private Sub openStrage()
        putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "Open the Root Dir")

        Process.Start(env.rootDir)
    End Sub

    ''' <summary>
    ''' データ格納DBの初期化
    ''' </summary>
    Private Sub initializeDB()
        Dim file_pass As String
        Dim sql As New List(Of String)
        Dim count As Integer = 0

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

        If loginDir(env.summaryDrive, env.summaryPath, env.loginId, env.loginPass, msg) Then
            dirInfo = New IO.DirectoryInfo(env.summaryPath)
            allFile = dirInfo.GetFiles(env.csvFile, IO.SearchOption.AllDirectories)
            For Each f As IO.FileInfo In allFile
                csvlist.Add(f.FullName)
            Next

            If csvlist.Count > 0 Then
                putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Get the CSV File count = " & csvlist.Count)
            Else
                putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  CAUTION : CSV File is Not Found")
            End If

            If Not logoffDir(env.summaryDrive, msg) Then
                putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  CAUTION : Network Dir logoff Error. Win32 API Error Code = " & msg & " [" & env.summaryPath & "]")
            End If

        Else
            putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  CAUTION : Network Dir login Error. Win32 API Error Code = " & msg & " [" & env.summaryPath & "]")
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
        Dim newPatId As String = ""
        Dim newFilePath As String = ""
        Dim newName As String = ""
        Dim strSql As New List(Of String)
        Dim strItem As String
        Dim index As Integer = 0
        Dim msg As String = ""

        putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "Regist File Data with DB")

        If loginDir(env.summaryDrive, env.summaryPath, env.loginId, env.loginPass, msg) Then
            For index = 0 To csvlist.Count - 1
                sRead = New IO.StreamReader(csvlist(index), System.Text.Encoding.GetEncoding("Shift_JIS"))

                Do
                    strline = sRead.ReadLine()
                    If strline Is Nothing Then
                        Exit Do
                    End If

                    outputFlg = " "
                    newPatId = " "
                    newFilePath = " "
                    newName = " "
                    checkMigration(strline, outputFlg, newPatId, newFilePath, newName)
                    strItem = strMid(csvlist.Item(index), 1, csvlist.Item(index).LastIndexOf("\"))
                    If outputFlg = "C" Then
                        strSql.Add("insert into SummaryData values (" & Chr(34) & strItem & Chr(34) & ", " & strline & ", " & Chr(34) & outputFlg & Chr(34) & ", " & Chr(34) & newPatId & Chr(34) & ", " & Chr(34) & newFilePath & Chr(34) & ", " & Chr(34) & newName & Chr(34) & ")")
                    End If
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

            If Not logoffDir(env.summaryDrive, msg) Then
                putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  CAUTION : Network Dir logoff Error. Win32 API Error Code = " & msg & " [" & env.summaryPath & "]")
            End If
        Else
            putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  CAUTION : Network Dir login Error. Win32 API Error Code = " & msg & " [" & env.summaryPath & "]")
        End If

    End Sub

    ''' <summary>
    ''' 移行対象かチェックして、移行用のファイル情報を作成
    ''' </summary>
    ''' <param name="strLine"></param>
    ''' <param name="outputFlg"></param>
    ''' <param name="newPadId"></param>
    ''' <param name="newName"></param>
    Private Sub checkMigration(strLine As String, ByRef outputFlg As String, ByRef newPadId As String, ByRef newFilePath As String, ByRef newName As String)
        Dim strLineAttr() As String
        Dim strTmp() As String
        Dim docCode As String = ""
        Dim docDate As String = ""
        Dim orderNo As String = ""
        Dim transactionDate As String = ""
        Dim deptCode As String = ""

        strLineAttr = Split(strLine, ",")
        docCode = strLineAttr(3).Replace(Chr(34), "")
        If env.getTargetCode().IndexOf(docCode) <> -1 Then
            strTmp = Split(strLineAttr(1), "_")
            'docDate = strTmp(1)
            docDate = strMid(strLineAttr(11).Replace(Chr(34), ""), 1, 8)    'カルテ移行ツールに合わせて、更新日をDocDateにする

            Try
                If Integer.Parse(docDate) >= Integer.Parse(env.fromDay) Then
                    If Integer.Parse(docDate) <= Integer.Parse(env.toDay) Then
                        newPadId = strLineAttr(2).Replace(Chr(34), "").PadLeft(env.pidLen, env.pidChr)
                        orderNo = strTmp(6).Replace(Chr(34), "").Replace(".PDF", "")
                        transactionDate = strLineAttr(11).Replace(Chr(34), "")
                        deptCode = strLineAttr(5).Replace(Chr(34), "")
                        outputFlg = "C"

                        newFilePath = strMid(newPadId, 1, 3) & "\" & strMid(newPadId, 4, 3) & "\" & newPadId & "\" & docDate & "\" & env.newDocCode
                        newName = newPadId & "_" & docDate & "_" & env.newDocCode & "_" & orderNo & "_" & transactionDate & "_" & deptCode & "_1.pdf"
                    End If
                End If
            Catch ex As Exception
                outputFlg = ""
            End Try
        End If
    End Sub

    ''' <summary>
    ''' PDFファイルをコピーする
    ''' </summary>
    ''' <returns></returns>
    Private Function copySummaryFile() As Integer
        Dim dtTbl As New DataTable
        Dim dataDir As String = ""
        Dim oldFile As String = ""
        Dim count As Integer = 0
        Dim msg As String = ""

        putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "Start SummaryFile Copy")

        If loginDir(env.summaryDrive, env.summaryPath, env.loginId, env.loginPass, msg) Then
            If csvData.ProcSelect("select * from SummaryData where outputFlg = 'C'", dtTbl) Then
                For Each tbldata As DataRow In dtTbl.Rows
                    oldFile = tbldata("CsvPath").ToString() & "\" & tbldata("DocID").ToString
                    If checkExists(oldFile, True) Then
                        dataDir = env.rootDir & "\" & tbldata("NewFilePath").ToString

                        If Not checkExists(dataDir, False) Then
                            IO.Directory.CreateDirectory(dataDir)
                        End If

                        IO.File.Copy(oldFile, dataDir & "\" & tbldata("NewFileName").ToString, True)

                        count = count + 1
                        If (count Mod 100) = 0 Then
                            putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  SummaryFile = " & count)
                        End If
                    Else
                        putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  CAUTION : Base File is not Found [" & oldFile & "]")
                    End If
                Next

                putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Total Copy SummaryFile = " & count)
            Else
                putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  CAUTION : SQL Select Error")
            End If
        Else
            putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  CAUTION : Network Dir login Error. Win32 API Error Code = " & msg & " [" & env.summaryPath & "]")
        End If

        Return count
    End Function

    ''' <summary>
    ''' コピーされているかチェックしてDBを更新
    ''' </summary>
    ''' <returns></returns>
    Private Function checkSummaryFile() As Integer
        Dim dtTbl As New DataTable
        Dim strSql As New List(Of String)
        Dim count As Integer = 0

        putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "SummaryFile Existence Check")

        If csvData.ProcSelect("select * from SummaryData where outputFlg = 'C'", dtTbl) Then
            For Each tblData As DataRow In dtTbl.Rows
                If checkExists(env.rootDir & "\" & tblData("NewFilePath").ToString & "\" & tblData("NewFileName").ToString, True) Then
                    strSql.Add("update SummaryData set OutputFlg = 'U' where NewFilePath = '" & tblData("NewFilePath").ToString & "' and NewFileName = '" & tblData("NewFileName").ToString & "';")
                End If
            Next

            If strSql.Count > 0 Then
                If csvData.ProcUpdate(strSql, count) Then
                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  Check data = " & count)
                Else
                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  CAUTION : Failed to update DB")
                End If
            Else
                putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  CAUTION : Data is not Found")
            End If
        Else
            putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "  CAUTION : SQL Select Error")
        End If

        Return count
    End Function

    ''' <summary>
    ''' データグリッドビューの内容をファイル出力
    ''' </summary>
    ''' <param name="dataview"></param>
    ''' <param name="file_pass"></param>
    ''' <returns></returns>
    Private Function saveCsv(dataview As DataGridView, ByRef file_pass As String) As Boolean
        Dim ret As Boolean = False
        Dim fileName As String = ""
        Dim result As New System.Text.StringBuilder
        Dim sWrite As IO.StreamWriter
        Dim i As Integer = 0
        Dim j As Integer = 0

        If dataview.Rows.Count = 0 Then
            Return ret
        End If

        If sfd.ShowDialog = Windows.Forms.DialogResult.OK Then
            For i = 0 To dataview.Columns.Count - 1
                Select Case i
                    Case 0
                        result.Append("""" & dataview.Columns(i).HeaderText.ToString & """")
                    Case dataview.Columns.Count - 1
                        result.Append("," & """" & dataview.Columns(i).HeaderText.ToString & """" & vbCrLf)
                    Case Else
                        result.Append("," & """" & dataview.Columns(i).HeaderText.ToString & """")
                End Select
            Next

            For i = 0 To dataview.Rows.Count - 1
                For j = 0 To dataview.Columns.Count - 1
                    Select Case j
                        Case 0
                            result.Append("""" & dataview.Rows(i).Cells(j).Value.ToString & """")
                        Case dataview.Columns.Count - 1
                            result.Append("," & """" & dataview.Rows(i).Cells(j).Value.ToString & """" & vbCrLf)
                        Case Else
                            result.Append("," & """" & dataview.Rows(i).Cells(j).Value.ToString & """")
                    End Select
                Next
            Next

            fileName = sfd.FileName
            sWrite = New IO.StreamWriter(fileName, False, System.Text.Encoding.GetEncoding("Shift_JIS"))
            sWrite.WriteLine(result)
            sWrite.Close()

            ret = True
        End If

        Return ret
    End Function

End Class
