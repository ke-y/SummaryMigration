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
        putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "Start System")

        Text = Application.ProductName & " - MAIN"

        'iniファイルの読み込み
        readIni()
        'DB初期化
        initialDB()

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
        Dim tmpStr As String

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

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "Set the Editor path")
                                Else
                                    errList.Add(Trim(strTmp(1)) & "が存在しません")
                                End If
                            Case "CSV_PATH"
                                If Trim(strTmp(1)) <> "" Then
                                    env.csvpath = Trim(strTmp(1))

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "Set the CSV input path")
                                Else
                                    errList.Add("CSV_PATHが未指定です")
                                End If
                            Case "CSV_ID"
                                If Trim(strTmp(1)) <> "" Then
                                    env.csvId = Trim(strTmp(1))

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "Set the CSV access ID")
                                End If
                            Case "CSV_PASS"
                                If Trim(strTmp(1)) <> "" Then
                                    env.csvPass = Trim(strTmp(1))

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "Set the CSV access Pass")
                                End If
                            Case "CSV_FILE"
                                If Trim(strTmp(1)) <> "" Then
                                    env.csvFile = Trim(strTmp(1))

                                    putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "Set the CSV file name pattern")
                                End If
                        End Select
                    End If
                End If
            Loop
            sRead.Close()
        Else
            MsgBox(file_pass & "が存在しません" & vbCrLf & "システムを終了します")

            putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "Stop System")
            Close()
        End If

        tmpStr = ""
        If loginDir(env.csvDrive, env.csvPath, env.csvId, env.csvPass, tmpStr) Then
            putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "Access OK " & " [" & env.csvPath & "]")

            If Not logoffDir(env.csvDrive, tmpStr) Then
                putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "CAUTION : Count Dir logoff Error. Win32 API Error Code = " & tmpStr & " [" & env.csvPath & "]")
            End If
        Else
            errList.Add(env.csvPath & "にアクセスできません")

            putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "CAUTION : Count Dir login Error. Win32 API Error Code = " & tmpStr & " [" & env.csvPath & "]")
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
    Private Sub initialDB()
        Dim file_pass As String
        Dim msg As String

        putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "Iinitialize DB")

        file_pass = env.appPath & env.dbPath
        If checkExists(file_pass, True) Then
            csvData.setPath = file_pass
            msg = csvData.initDB()

            putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", msg)
        Else
            MsgBox(file_pass & "が存在しません" & vbCrLf & "システムを終了します")

            putLog(env.appPath & "\" & env.appLog, My.Application.Info.ProductName & "_" & Date.Now.ToString("yyyyMMdd") & ".log", "Stop System")
            Close()
        End If

    End Sub
End Class
