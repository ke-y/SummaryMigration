Option Explicit On

Public Class frmMain
    Private env As New envData

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

        If errList.Count > 0 Then
            For i = 0 To errList.Count - 1
                MsgBox(errList.Item(i))
            Next i
        End If

    End Sub

End Class
